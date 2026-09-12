"""PPTX -> positioned HTML/SVG, using Python's standard library only.
Deliberately bounded: unsupported features are explicit, never silently dropped.
"""
from __future__ import annotations
import base64
import colorsys
import hashlib
import html
import json
import math
import posixpath
import re
from collections import Counter
from pathlib import Path
from urllib.parse import urlsplit
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
EMU = 9525.0  # 914400 EMU/in at 96 CSS px/in; 1pt = 4/3 CSS px
SAFE_IMAGES = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
               '.gif': 'image/gif', '.webp': 'image/webp', '.svg': 'image/svg+xml', '.bmp': 'image/bmp'}
PRESETS = {'rect', 'roundRect', 'ellipse', 'line', 'triangle', 'rtTriangle', 'diamond',
           'parallelogram', 'trapezoid', 'hexagon', 'pentagon', 'chevron', 'rightArrow', 'leftArrow'}

def tag(n): return n.tag.rsplit('}', 1)[-1]
def find(n, p): return n.find(p, NS) if n is not None else None
def allof(n, p): return n.findall(p, NS) if n is not None else []
def attr(n, k, default=None): return n.get(k, default) if n is not None else default
def num(v, default=0.0):
    try:
        value=float(v)
        return value if math.isfinite(value) else default
    except (TypeError, ValueError): return default

def px(v): return num(v) / EMU
def nfmt(v): return f'{float(v):.4f}'.rstrip('0').rstrip('.') or '0'
def esc(v): return html.escape(str(v), quote=True)
def safe_json(v): return json.dumps(v, ensure_ascii=False).replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026')
def css_family(v): return '"' + str(v).replace('\\', '\\\\').replace('"', '\\"').replace('<', '\\3C ').replace('>', '\\3E ').replace('\n', ' ').replace('\r', ' ') + '"'
def data_uri(data, mime): return 'data:' + mime + ';base64,' + base64.b64encode(data).decode('ascii')
def first(nodes): return next((n for n in nodes if n is not None), None)
def chosen(nodes, path): return first(find(n, path) for n in nodes)
def avalue(nodes, key, default=None): return next((n.get(key) for n in nodes if n is not None and key in n.attrib), default)

class Converter:
    def __init__(self, source, include_notes=False, fallbacks=None):
        self.source = Path(source)
        self.z = ZipFile(self.source)
        entries = self.z.infolist()
        if len(entries) > 20000 or sum(i.file_size for i in entries) > 512 * 1024 * 1024:
            raise ValueError('PPTX package exceeds safe processing limits')
        if self.z.testzip(): raise ValueError('Corrupt PPTX ZIP package')
        self.cache, self.relcache = {}, {}
        self.issues, self.fonts, self.texts, self.stats = [], set(), [], Counter()
        self.include_notes, self.fallbacks = include_notes, fallbacks or {}
        self.used_fallbacks=set()
        self.presentation = self.xml('ppt/presentation.xml')
        sz = find(self.presentation, 'p:sldSz')
        self.width, self.height = px(attr(sz, 'cx')), px(attr(sz, 'cy'))
        if self.width <= 0 or self.height <= 0: raise ValueError('Invalid slide dimensions')
        self.slide_no, self.oid, self.part = 0, '', ''
        self.layout, self.master, self.theme = None, None, None
        self.palette, self.clrmap = {}, {}
        self.default_text = find(self.presentation, 'p:defaultTextStyle')
        self.notes = []

    def close(self): self.z.close()

    def xml(self, name):
        if name not in self.cache:
            if name not in self.z.namelist(): return None
            raw = self.z.read(name)
            if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():
                raise ValueError('DTD/entities are not accepted in OOXML')
            self.cache[name] = ET.fromstring(raw)
        return self.cache[name]

    def rels(self, part):
        if part not in self.relcache:
            path = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
            root = self.xml(path)
            rels = {}
            for r in root if root is not None else []:
                external = r.get('TargetMode') == 'External'
                target = r.get('Target', '')
                if not external:
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(part), target)) if not target.startswith('/') else target.lstrip('/')
                    if target.startswith('../'): raise ValueError('Invalid package relationship')
                rels[r.get('Id')] = {'target': target, 'external': external, 'type': r.get('Type', '').rsplit('/', 1)[-1]}
            self.relcache[part] = rels
        return self.relcache[part]

    def related(self, part, kind):
        return next((r['target'] for r in self.rels(part).values() if r['type'] == kind and not r['external']), '')

    def issue(self, feature, message, severity='error'):
        item = dict(slide=self.slide_no, object_id=self.oid, feature=feature, severity=severity, message=message)
        if item not in self.issues: self.issues.append(item)

    def color(self, node, default='#000000'):
        if node is None: return default
        child = node if tag(node).endswith('Clr') else next((c for c in node if tag(c).endswith('Clr')), None)
        if child is None: return default
        kind, value = tag(child), child.get('val', '')
        if kind == 'srgbClr': base = '#' + value
        elif kind == 'sysClr': base = '#' + child.get('lastClr', '000000')
        elif kind == 'schemeClr': base = self.palette.get(self.clrmap.get(value, value), default)
        elif kind == 'scrgbClr': base = '#' + ''.join(f'{round(max(0,min(100000,num(child.get(k))))*255/100000):02X}' for k in ('r','g','b'))
        elif kind == 'prstClr':
            named = {'black':'#000000','white':'#FFFFFF','red':'#FF0000','blue':'#0000FF','green':'#008000','yellow':'#FFFF00','gray':'#808080'}
            base = named.get(value, default)
            if value not in named: self.issue('preset-color', 'Named preset color requires review', 'warning')
        else:
            self.issue('color', 'Unsupported color representation')
            return default
        if not re.fullmatch(r'#[0-9a-fA-F]{6}', base): base = default
        rgb = [int(base[i:i+2], 16)/255 for i in (1,3,5)]
        alpha = 1.0
        for tr in child:
            key, v = tag(tr), num(tr.get('val')) / 100000
            if key == 'alpha': alpha = v
            elif key == 'alphaMod': alpha *= v
            elif key in ('lumMod','lumOff','satMod','satOff'):
                h,l,s = colorsys.rgb_to_hls(*rgb)
                if key == 'lumMod': l *= v
                elif key == 'lumOff': l += v
                elif key == 'satMod': s *= v
                else: s += v
                rgb = list(colorsys.hls_to_rgb(h,max(0,min(1,l)),max(0,min(1,s))))
            elif key == 'tint': rgb = [x + (1-x)*v for x in rgb]
            elif key == 'shade': rgb = [x*v for x in rgb]
            else: self.issue('color-transform', 'Unsupported color transform: '+key, 'warning')
        channels = [round(max(0,min(1,x))*255) for x in rgb]
        return f'rgba({channels[0]},{channels[1]},{channels[2]},{nfmt(alpha)})' if alpha != 1 else '#' + ''.join(f'{v:02X}' for v in channels)

    def setup_theme(self, master_part, slide, layout):
        self.theme = self.xml(self.related(master_part, 'theme'))
        self.palette = {}
        scheme = find(self.theme, 'a:themeElements/a:clrScheme')
        for c in scheme if scheme is not None else []: self.palette[tag(c)] = self.color(c)
        cm = find(self.master, 'p:clrMap')
        self.clrmap = dict(cm.attrib) if cm is not None else {'tx1':'dk1','bg1':'lt1','tx2':'dk2','bg2':'lt2'}
        for root in (layout, slide):
            override = find(root, 'p:clrMapOvr/a:overrideClrMapping')
            if override is not None: self.clrmap.update(override.attrib)

    def fill(self, nodes, allow_style=False):
        for node in nodes:
            if node is None: continue
            if find(node, 'a:noFill') is not None: return 'none', ''
            solid = find(node, 'a:solidFill')
            if solid is not None: return self.color(solid), ''
            grad = find(node, 'a:gradFill')
            if grad is not None:
                linear = find(grad, 'a:lin')
                if linear is None:
                    self.issue('gradient', 'Only linear gradients are supported')
                    return '#D1D5DB', ''
                angle = num(linear.get('ang'))/60000 * math.pi/180
                c,s = math.cos(angle), math.sin(angle)
                gid = self.oid + '-gradient-' + str(self.stats['gradients'])
                self.stats['gradients'] += 1
                stops = ''.join(f'<stop offset="{nfmt(num(g.get("pos"))/1000)}%" stop-color="{esc(self.color(g))}"/>' for g in allof(grad,'a:gsLst/a:gs'))
                defs = f'<linearGradient id="{gid}" x1="{nfmt(50-50*c)}%" y1="{nfmt(50-50*s)}%" x2="{nfmt(50+50*c)}%" y2="{nfmt(50+50*s)}%">{stops}</linearGradient>'
                return f'url(#{gid})', defs
            if any(find(node, 'a:'+k) is not None for k in ('blipFill','pattFill','grpFill')):
                self.issue('fill', 'Picture, pattern or inherited group fill requires a fallback')
                return '#E5E7EB', ''
        if allow_style:
            self.issue('theme-fill', 'Implicit theme fill is approximated; verify or supply an element fallback', 'warning')
        return 'none', ''

    def reference_style(self, chain, refname, listname):
        ref=chosen(chain,'p:style/a:'+refname)
        index=int(num(attr(ref,'idx')))
        if index<=0: return None
        if index>=1001 and listname=='fillStyleLst': listname='bgFillStyleLst'; index-=1000
        container=find(self.theme,'a:themeElements/a:fmtScheme/a:'+listname)
        if container is None or index>len(container):
            self.issue('theme-style','Missing referenced theme style','warning'); return None
        self.palette['phClr']=self.color(ref)
        return container[index-1]

    def shadow(self, effects):
        result=''
        for effect in effects if effects is not None else []:
            if tag(effect)!='outerShdw':
                self.issue('shape-effects','Only outer shadow is approximated; other effects need fallback'); continue
            if any(effect.get(k) not in (None,v) for k,v in [('sx','100000'),('sy','100000'),('kx','0'),('ky','0')]):
                self.issue('shadow-transform','Transformed shadow requires fallback'); continue
            angle=num(effect.get('dir'))/60000*math.pi/180
            dist=px(effect.get('dist')); blur=px(effect.get('blurRad'))/2
            result+=f' drop-shadow({nfmt(math.cos(angle)*dist)}px {nfmt(math.sin(angle)*dist)}px {nfmt(blur)}px {self.color(effect)})'
            self.stats['outer_shadows']+=1
        return 'filter:'+result.strip()+';' if result else ''

    def stroke(self, node):
        if node is None: return 'stroke="none"', ''
        paint, defs = self.fill([node])
        if paint == 'none': return 'stroke="none"', defs
        width = px(node.get('w', '12700'))
        attrs = f'stroke="{esc(paint)}" stroke-width="{nfmt(width)}"'
        dash = attr(find(node,'a:prstDash'), 'val', 'solid')
        patterns = {'dash':'4 3','dot':'1 2','dashDot':'4 2 1 2','lgDash':'8 3','sysDot':'1 1','sysDash':'3 1'}
        if dash in patterns: attrs += f' stroke-dasharray="{patterns[dash]}"'
        elif dash != 'solid': self.issue('line-dash', 'Unsupported dash preset', 'warning')
        if node.get('cmpd', 'sng') != 'sng': self.issue('compound-line', 'Compound line is approximated', 'warning')
        for path, side in [('a:headEnd','start'),('a:tailEnd','end')]:
            end = find(node,path)
            endtype = attr(end,'type','none')
            if endtype != 'none':
                mid = self.oid + '-arrow-' + side
                if endtype not in ('triangle','arrow','stealth'): self.issue('arrowhead', 'Arrowhead preset approximated', 'warning')
                d = 'M0,0 L5,2.5 L0,5' + ('' if endtype == 'arrow' else ' Z')
                marker = f'<marker id="{mid}" viewBox="0 0 5 5" refX="4.5" refY="2.5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="{d}" fill="{esc(paint) if endtype != "arrow" else "none"}" stroke="{esc(paint)}" stroke-width="0.6"/></marker>'
                defs += marker
                attrs += f' marker-{side}="url(#{mid})"'
        return attrs, defs

    def ph(self, shape): return find(shape,'.//p:nvPr/p:ph')

    def placeholder(self, root, ph, master=False):
        if ph is None or root is None: return None
        choices = allof(root,'p:cSld/p:spTree/p:sp')
        for s in choices:
            other = self.ph(s)
            if other is None: continue
            if master:
                if other.get('type','body') == ph.get('type','body'): return s
            elif other.get('idx','0') == ph.get('idx','0'): return s
        return None

    def chain(self, shape, origin):
        nodes = [shape]
        ph = self.ph(shape)
        if ph is not None and origin == 'slide':
            parent = self.placeholder(self.layout, ph)
            if parent is not None: nodes.append(parent)
            parent = self.placeholder(self.master, ph, True)
            if parent is not None: nodes.append(parent)
        return nodes

    def font(self, node, defaults):
        props = [node] + defaults
        family_node = chosen(props,'a:ea')
        if not attr(family_node,'typeface'): family_node = chosen(props,'a:latin')
        family = attr(family_node,'typeface','Arial') or 'Arial'
        if family.startswith('+'):
            branch = 'majorFont' if family.startswith('+mj') else 'minorFont'
            path = 'a:themeElements/a:fontScheme/a:' + branch
            family = attr(find(self.theme,path+'/a:ea'),'typeface') or attr(find(self.theme,path+'/a:latin'),'typeface') or 'Arial'
        self.fonts.add(family)
        size = num(avalue(props,'sz','1800'))/100 * 4/3
        color = self.color(chosen(props,'a:solidFill'))
        style = f'font-family:{css_family(family)},sans-serif;font-size:{nfmt(size)}px;font-weight:{"700" if avalue(props,"b","0")=="1" else "400"};color:{color};'
        if avalue(props,'i','0') == '1': style += 'font-style:italic;'
        decorations=[]
        if avalue(props,'u','none') != 'none': decorations.append('underline')
        if avalue(props,'strike','noStrike') != 'noStrike': decorations.append('line-through')
        if decorations: style += 'text-decoration:'+' '.join(decorations)+';'
        spacing = avalue(props,'spc')
        if spacing is not None: style += f'letter-spacing:{nfmt(num(spacing)/100*4/3)}px;'
        base = avalue(props,'baseline')
        if base and num(base):
            style += f'vertical-align:{nfmt(num(base)/1000)}%;'
            self.issue('baseline', 'Superscript/subscript requires visual comparison', 'warning')
        if chosen(props,'a:effectLst') is not None and len(chosen(props,'a:effectLst')): self.issue('text-effects','Text effects require fallback')
        return style, size

    def text(self, tx, ancestors, ph_type='other', tcpr=None):
        if tx is None or (not allof(tx,'.//a:t') and not allof(tx,'.//a:br')): return ''
        bodies = [tx] + [find(n,'p:txBody') for n in ancestors]
        bodyprops = [find(n,'a:bodyPr') for n in bodies]
        # Table cell inset defaults differ from text-frame defaults.
        if tcpr is not None:
            margins = [px(tcpr.get(k,d)) for k,d in [('marT','45720'),('marR','91440'),('marB','45720'),('marL','91440')]]
            anchor = tcpr.get('anchor', avalue(bodyprops,'anchor','t'))
        else:
            margins = [px(avalue(bodyprops,k,d)) for k,d in [('tIns','45720'),('rIns','91440'),('bIns','45720'),('lIns','91440')]]
            anchor = avalue(bodyprops,'anchor','t')
        if avalue(bodyprops,'vert','horz') != 'horz': self.issue('vertical-text','Vertical/WordArt text is not supported')
        if num(avalue(bodyprops,'numCol','1')) > 1: self.issue('text-columns','Multi-column text requires fallback')
        if num(avalue(bodyprops,'rot','0')): self.issue('text-rotation','Independent text rotation is not supported')
        if chosen(bodyprops,'a:prstTxWarp') is not None: self.issue('wordart','Warped text requires fallback')
        autofit = chosen(bodyprops,'a:normAutofit')
        if autofit is not None and num(autofit.get('fontScale','100000')) != 100000:
            self.issue('autofit','Stored fontScale not applied; inspect or supply fallback')
        style = 'position:absolute;inset:0;box-sizing:border-box;display:flex;flex-direction:column;'
        style += 'justify-content:' + {'t':'flex-start','ctr':'center','b':'flex-end'}.get(anchor,'flex-start') + ';'
        if anchor not in ('t','ctr','b'): self.issue('text-anchor','Distributed vertical alignment approximated','warning')
        style += 'padding:'+' '.join(nfmt(m)+'px' for m in margins)+';'
        nowrap = avalue(bodyprops,'wrap','square') == 'none'
        style += 'white-space:' + ('pre' if nowrap else 'pre-wrap') + ';overflow:visible;word-break:normal;overflow-wrap:normal;'
        paragraphs = []
        master_style = {'title':'titleStyle','ctrTitle':'titleStyle','body':'bodyStyle'}.get(ph_type,'otherStyle')
        for pi,p in enumerate(allof(tx,'a:p')):
            ppr = find(p,'a:pPr')
            level = int(num(attr(ppr,'lvl','0'))) + 1
            style_path = f'a:lvl{level}pPr'
            pprops = [ppr,find(tx,'a:lstStyle/'+style_path)]
            for body in bodies[1:]:
                pprops += [find(body,'a:p/a:pPr'),find(body,'a:lstStyle/'+style_path)]
            pprops += [find(self.master,'p:txStyles/p:'+master_style+'/'+style_path),find(self.default_text,style_path)]
            defaults = [find(pp,'a:defRPr') for pp in pprops]
            first_run_pr = first(find(r,'a:rPr') for r in p if tag(r) in ('r','fld'))
            base_style, base_size = self.font(first_run_pr if first_run_pr is not None else find(p,'a:endParaRPr'),defaults)
            align = avalue(pprops,'algn','l')
            ps = 'margin:0;flex-shrink:0;text-align:'+{'l':'left','r':'right','ctr':'center','just':'justify','dist':'justify'}.get(align,'left')+';'
            ps += 'margin-left:'+nfmt(px(avalue(pprops,'marL','0')))+'px;text-indent:'+nfmt(px(avalue(pprops,'indent','0')))+'px;'
            for child,css in [('a:spcBef','margin-top'),('a:spcAft','margin-bottom')]:
                sp = chosen(pprops,child)
                pts = find(sp,'a:spcPts')
                pct = find(sp,'a:spcPct')
                val = num(attr(pts,'val'))/100*4/3 if pts is not None else base_size*num(attr(pct,'val'))/100000
                ps += css+':'+nfmt(val)+'px;'
            sp = chosen(pprops,'a:lnSpc')
            pts,pct = find(sp,'a:spcPts'),find(sp,'a:spcPct')
            line = nfmt(num(pts.get('val'))/100*4/3)+'px' if pts is not None else nfmt(num(attr(pct,'val','100000'))/100000)
            ps += 'line-height:'+line+';'+base_style
            bullet = next((c for pp in pprops if pp is not None for c in pp if tag(c) in ('buNone','buChar','buAutoNum','buBlip')),None)
            runs = []
            if bullet is not None and tag(bullet) != 'buNone':
                if tag(bullet) == 'buChar': marker=bullet.get('char','•')
                elif tag(bullet) == 'buAutoNum':
                    marker=str(pi+int(bullet.get('startAt','1')))+'.'
                    self.issue('numbered-list','Numbering style approximated as decimal', 'warning')
                else:
                    marker=''; self.issue('picture-bullet','Picture bullets not supported')
                runs.append('<span class="ppt-bullet" aria-hidden="true">'+esc(marker)+' </span>')
            for run in p:
                kind = tag(run)
                if kind == 'br': runs.append('<br>'); continue
                if kind not in ('r','fld'): continue
                text_node = find(run,'a:t')
                value = (text_node.text or '') if text_node is not None else ''
                if kind == 'fld' and run.get('type') == 'slidenum': value = str(self.slide_no)
                elif kind == 'fld': self.issue('field','Dynamic field uses stored display text','warning')
                self.texts.append(value)
                runpr=find(run,'a:rPr')
                rs,_ = self.font(runpr,defaults)
                inside=esc(value)
                link=find(runpr,'a:hlinkClick')
                if link is not None:
                    rel=self.rels(self.part).get(link.get('{'+NS['r']+'}id'),{})
                    url=rel.get('target','')
                    if rel.get('external') and urlsplit(url).scheme.lower() in ('https','http','mailto','tel'):
                        inside='<a href="'+esc(url)+'" target="_blank" rel="noopener noreferrer">'+inside+'</a>'
                    else: self.issue('hyperlink','Unsafe or internal/action hyperlink not converted','warning')
                runs.append('<span class="ppt-run" data-editable="true" style="'+esc(rs)+'">'+inside+'</span>')
            if not runs: runs=['<br>']
            paragraphs.append('<p style="'+esc(ps)+'">'+''.join(runs)+'</p>')
        # Empty auto-shapes need no phantom text boxes or font/overflow warnings.
        if not allof(tx,'.//a:t') and not allof(tx,'.//a:br'): return ''
        self.stats['text_frames'] += 1
        return '<div class="ppt-text" style="'+esc(style)+'">'+''.join(paragraphs)+'</div>'

    def image(self, data, suffix):
        mime=SAFE_IMAGES.get(suffix.lower())
        if not mime: raise ValueError('Unsupported image format: '+suffix)
        if mime == 'image/svg+xml':
            if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper(): raise ValueError('Unsafe SVG entities')
            root=ET.fromstring(data)
            for el in root.iter():
                if tag(el) in ('script','foreignObject','style','animate','set'): raise ValueError('Active SVG is not accepted')
                for key,val in el.attrib.items():
                    if key.lower().startswith('on') or (key.endswith('href') and not val.startswith('#')) or ('url(' in val and not re.fullmatch(r'url\(#[\w-]+\)',val)):
                        raise ValueError('SVG active/external resource is not accepted')
        return data_uri(data,mime)

    def geometry(self, geom, w, h):
        if geom is None: return ''
        if tag(geom) == 'custGeom':
            paths=[]
            for path in allof(geom,'a:pathLst/a:path'):
                sx,sy=w/max(1,num(path.get('w'),w*EMU)),h/max(1,num(path.get('h'),h*EMU))
                commands=[]
                for cmd in path:
                    k=tag(cmd)
                    code={'moveTo':'M','lnTo':'L','cubicBezTo':'C','quadBezTo':'Q','close':'Z'}.get(k)
                    if not code:
                        self.issue('custom-geometry','Unsupported custom path command '+k); continue
                    pts=[]
                    for pt in allof(cmd,'a:pt'):
                        if not re.fullmatch(r'-?\d+(\.\d+)?',pt.get('x','')) or not re.fullmatch(r'-?\d+(\.\d+)?',pt.get('y','')):
                            self.issue('custom-geometry','Formula path coordinates require fallback')
                        pts.append(nfmt(num(pt.get('x'))*sx)+','+nfmt(num(pt.get('y'))*sy))
                    commands.append(code+' '.join(pts))
                flags=' fill="none"' if path.get('fill') == 'none' else ''
                if path.get('stroke') == '0': flags+=' stroke="none"'
                paths.append('<path d="'+' '.join(commands)+'"'+flags+'/>')
            return ''.join(paths)
        preset=geom.get('prst','rect')
        if preset not in PRESETS:
            self.issue('preset-geometry','Unsupported preset '+preset)
            return '<rect width="'+nfmt(w)+'" height="'+nfmt(h)+'"/>'
        if preset == 'line': return f'<path fill="none" d="M0,0 L{nfmt(w)},{nfmt(h)}"/>'
        if preset in ('rect','roundRect'):
            adjust=attr(find(geom,'a:avLst/a:gd'),'fmla','val 16667').split()[-1]
            r=min(w,h)*num(adjust,16667)/100000 if preset=='roundRect' else 0
            return f'<rect width="{nfmt(w)}" height="{nfmt(h)}" rx="{nfmt(r)}"/>'
        if preset=='ellipse': return f'<ellipse cx="{nfmt(w/2)}" cy="{nfmt(h/2)}" rx="{nfmt(w/2)}" ry="{nfmt(h/2)}"/>'
        shapes={'triangle':[(.5,0),(1,1),(0,1)],'rtTriangle':[(0,0),(1,1),(0,1)],'diamond':[(.5,0),(1,.5),(.5,1),(0,.5)],
                'parallelogram':[(.2,0),(1,0),(.8,1),(0,1)],'trapezoid':[(.2,0),(.8,0),(1,1),(0,1)],
                'hexagon':[(.25,0),(.75,0),(1,.5),(.75,1),(.25,1),(0,.5)],'pentagon':[(.5,0),(1,.38),(.81,1),(.19,1),(0,.38)],
                'chevron':[(0,0),(.7,0),(1,.5),(.7,1),(0,1),(.3,.5)],
                'rightArrow':[(0,.25),(.65,.25),(.65,0),(1,.5),(.65,1),(.65,.75),(0,.75)],
                'leftArrow':[(1,.25),(.35,.25),(.35,0),(0,.5),(.35,1),(.35,.75),(1,.75)]}
        if allof(geom,'a:avLst/a:gd'): self.issue('geometry-adjustments','Non-rect preset adjustment uses default outline','warning')
        return '<polygon points="'+' '.join(nfmt(x*w)+','+nfmt(y*h) for x,y in shapes[preset])+'"/>'

    def table(self, table, w, h):
        cols=[px(c.get('w')) for c in allof(table,'a:tblGrid/a:gridCol')]
        rows=allof(table,'a:tr')
        scale_x=w/sum(cols) if sum(cols) else 1
        heights=[px(r.get('h')) for r in rows]
        scale_y=h/sum(heights) if sum(heights) else 1
        out='<table class="ppt-table" style="width:100%;height:100%;table-layout:fixed;border-collapse:collapse;border-spacing:0"><colgroup>'
        out+=''.join('<col style="width:'+nfmt(c/sum(cols)*100)+'%">' for c in cols)+'</colgroup><tbody>'
        for ri,row in enumerate(rows):
            out+='<tr style="height:'+nfmt(heights[ri]*scale_y)+'px">'
            for ci,cell in enumerate(allof(row,'a:tc')):
                if cell.get('hMerge')=='1' or cell.get('vMerge')=='1': continue
                pr=find(cell,'a:tcPr')
                fill,defs=self.fill([pr])
                if defs: self.issue('table-fill','Table cell gradient is not supported')
                if fill=='none' and find(pr,'a:noFill') is None:
                    self.issue('table-style','Implicit table-style fill is not resolved; use explicit cell formatting')
                rs=int(cell.get('rowSpan','1')); cs=int(cell.get('gridSpan','1'))
                cellh=sum(heights[ri:ri+rs])*scale_y
                st='position:relative;padding:0;vertical-align:top;height:'+nfmt(cellh)+'px;background:'+(fill if fill!='none' else 'transparent')+';'
                for k,css in [('lnL','left'),('lnR','right'),('lnT','top'),('lnB','bottom')]:
                    ln=find(pr,'a:'+k)
                    paint,_=self.fill([ln]); wid=px(attr(ln,'w','0'))
                    st+='border-'+css+':'+nfmt(wid)+'px solid '+(paint if paint!='none' else 'transparent')+';'
                out+=f'<td rowspan="{rs}" colspan="{cs}" style="{esc(st)}"><div style="position:absolute;inset:0">'
                out+=self.text(find(cell,'a:txBody'),[],tcpr=pr)+'</div></td>'
            out+='</tr>'
        self.stats['tables']+=1
        return out+'</tbody></table>'

    def render_shape(self, shape, origin, part, matrix=(1/EMU,1/EMU,0,0)):
        self.part=part
        nv=find(shape,'.//p:cNvPr')
        if attr(nv,'hidden','0')=='1': return ''
        sid=attr(nv,'id',str(self.stats['objects']+1))
        self.oid=f's{self.slide_no:03d}_{origin}_{sid}'
        oid=self.oid
        self.stats['objects']+=1
        name=attr(nv,'name','')
        kind=tag(shape)
        chain=self.chain(shape,origin)
        xfrm=chosen(chain,'p:spPr/a:xfrm')
        if kind=='graphicFrame': xfrm=find(shape,'p:xfrm')
        if kind=='grpSp': xfrm=find(shape,'p:grpSpPr/a:xfrm')
        if xfrm is None:
            self.issue('geometry','Missing shape transform')
        off,ext=find(xfrm,'a:off'),find(xfrm,'a:ext')
        sx,sy,tx,ty=matrix
        x=num(attr(off,'x'))*sx+tx; y=num(attr(off,'y'))*sy+ty
        w=num(attr(ext,'cx'))*sx; h=num(attr(ext,'cy'))*sy
        rotation=num(attr(xfrm,'rot'))/60000
        flipx,flipy=attr(xfrm,'flipH','0')=='1',attr(xfrm,'flipV','0')=='1'
        style=f'position:absolute;left:{nfmt(x)}px;top:{nfmt(y)}px;width:{nfmt(w)}px;height:{nfmt(h)}px;'
        if rotation: style+=f'transform:rotate({nfmt(rotation)}deg);transform-origin:center;'
        begin=f'<div class="ppt-object" id="{oid}" data-object-id="{oid}" data-kind="{kind}" data-name="{esc(name)}" style="{esc(style)}">'
        if oid in self.fallbacks:
            f=Path(self.fallbacks[oid]); uri=self.image(f.read_bytes(),f.suffix)
            self.stats['fallbacks']+=1
            self.used_fallbacks.add(oid)
            self.issue('element-fallback','Element raster/vector fallback used; text is not editable','warning')
            return begin+f'<img alt="{esc(name)}" src="{uri}" style="width:100%;height:100%"></div>'
        if kind=='grpSp':
            choff,chext=find(xfrm,'a:chOff'),find(xfrm,'a:chExt')
            nsx=w/max(1,num(attr(chext,'cx'))); nsy=h/max(1,num(attr(chext,'cy')))
            childmatrix=(nsx,nsy,-num(attr(choff,'x'))*nsx,-num(attr(choff,'y'))*nsy)
            if abs(nsx-1/EMU)>1e-8 or abs(nsy-1/EMU)>1e-8: self.issue('group-scale','Scaled group geometry is mapped; text and margin scaling requires visual review','warning')
            if flipx or flipy: self.issue('group-flip','Flipped groups require fallback')
            content=''.join(self.render_shape(c,origin,part,childmatrix) for c in shape if tag(c) in ('sp','pic','cxnSp','grpSp','graphicFrame'))
            self.stats['groups']+=1
            return begin+content+'</div>'
        content=''
        if kind=='graphicFrame':
            table=find(shape,'a:graphic/a:graphicData/a:tbl')
            if table is not None: content=self.table(table,w,h)
            else:
                gd=find(shape,'a:graphic/a:graphicData')
                feature=attr(gd,'uri','graphic').rsplit('/',1)[-1]
                self.issue(feature,'Native chart/SmartArt/OLE/graphic requires an element fallback or dedicated adapter')
                content='<div class="ppt-unsupported" style="position:absolute;inset:0;border:2px dashed #B91C1C;color:#B91C1C;background:#FFF1F2;padding:8px;font:14px sans-serif">검토 필요: '+esc(feature)+'</div>'
        elif kind=='pic':
            blip=find(shape,'p:blipFill/a:blip')
            picture_geom=find(shape,'p:spPr/a:prstGeom')
            if attr(picture_geom,'prst','rect')!='rect' or find(shape,'p:spPr/a:custGeom') is not None: self.issue('picture-mask','Non-rectangular image mask requires fallback')
            if find(shape,'p:blipFill/a:tile') is not None: self.issue('picture-tiling','Tiled image requires fallback')
            if any(tag(c)!='extLst' for c in (blip if blip is not None else [])): self.issue('picture-effects','Image color/transparency effects require fallback')
            pline=find(shape,'p:spPr/a:ln')
            if pline is not None and find(pline,'a:noFill') is None: self.issue('picture-border','Picture border requires fallback')
            rel=self.rels(part).get(attr(blip,'{'+NS['r']+'}embed'),{})
            if not rel or rel.get('external'):
                self.issue('linked-image','External image not fetched; embed it or supply fallback')
            else:
                try:
                    uri=self.image(self.z.read(rel['target']),Path(rel['target']).suffix)
                    crop=find(shape,'p:blipFill/a:srcRect')
                    l,t,r,b=[num(attr(crop,k))/100000 for k in ('l','t','r','b')]
                    fx,fy=1-l-r,1-t-b
                    if fx<=0 or fy<=0: raise ValueError('Invalid picture crop')
                    ist=f'position:absolute;left:{nfmt(-l/fx*100)}%;top:{nfmt(-t/fy*100)}%;width:{nfmt(100/fx)}%;height:{nfmt(100/fy)}%;max-width:none;'
                    content=f'<div style="position:absolute;inset:0;overflow:hidden;transform:scale({-1 if flipx else 1},{-1 if flipy else 1})"><img alt="{esc(attr(nv,"descr",name))}" src="{uri}" style="{ist}"></div>'
                    self.stats['images']+=1
                except (ValueError,KeyError,ET.ParseError) as exc: self.issue('image',str(exc))
            if find(shape,'.//a:videoFile') is not None or find(shape,'.//a:audioFile') is not None: self.issue('media','Video/audio is not converted')
        elif kind in ('sp','cxnSp'):
            props=[find(n,'p:spPr') for n in chain]
            geom=chosen(props,'a:custGeom')
            if geom is None: geom=chosen(props,'a:prstGeom')
            if geom is not None:
                fillprops=list(props)
                if not any(find(p,'a:'+k) is not None for p in props for k in ('solidFill','noFill','gradFill','blipFill','pattFill','grpFill')):
                    inherited=self.reference_style(chain,'fillRef','fillStyleLst')
                    if inherited is not None:
                        wrap=ET.Element('style'); wrap.append(inherited); fillprops.append(wrap)
                paint,defs=self.fill(fillprops)
                ln=chosen(props,'a:ln')
                if ln is None: ln=self.reference_style(chain,'lnRef','lnStyleLst')
                stroke,sd=self.stroke(ln)
                effects=chosen(props,'a:effectLst')
                if effects is None:
                    inherited=self.reference_style(chain,'effectRef','effectStyleLst')
                    effects=find(inherited,'a:effectLst')
                    if find(inherited,'a:scene3d') is not None or find(inherited,'a:sp3d') is not None: self.issue('3d','Inherited 3D effects require fallback')
                shadowcss=self.shadow(effects)
                defs+=sd
                drawing=self.geometry(geom,w,h)
                tf=''
                if flipx or flipy: tf=f' transform="translate({nfmt(w if flipx else 0)} {nfmt(h if flipy else 0)}) scale({-1 if flipx else 1} {-1 if flipy else 1})"'
                content+=f'<svg class="ppt-svg" aria-hidden="true" width="{nfmt(max(w,1))}" height="{nfmt(max(h,1))}" viewBox="0 0 {nfmt(max(w,1))} {nfmt(max(h,1))}" style="position:absolute;left:0;top:0;overflow:visible;pointer-events:none;{esc(shadowcss)}"><defs>{defs}</defs><g fill="{esc(paint)}" {stroke}{tf}>{drawing}</g></svg>'
                self.stats['svg_shapes']+=1
            txbody=find(shape,'p:txBody')
            ph=self.ph(shape)
            content+=self.text(txbody,chain[1:],attr(ph,'type','other'))
            for prop in props[:1]:
                if find(prop,'a:scene3d') is not None or find(prop,'a:sp3d') is not None: self.issue('3d','3D shapes not supported')
        else:
            self.issue('shape','Unsupported element '+kind)
        if not content and any(i['object_id']==oid and i['severity']=='error' for i in self.issues):
            content='<div style="border:2px dashed #B91C1C;color:#B91C1C;font:14px sans-serif">검토 필요: '+esc(kind)+'</div>'
        return begin+content+'</div>'

    def background(self, roots):
        bg=first(find(r,'p:cSld/p:bg/p:bgPr') for r in roots)
        if bg is None:
            if any(find(r,'p:cSld/p:bg/p:bgRef') is not None for r in roots): self.issue('background','Theme background reference requires review','warning')
            return 'background:#FFFFFF;', ''
        paint,defs=self.fill([bg])
        if defs:
            svg=f'<svg aria-hidden="true" style="position:absolute;inset:0;width:100%;height:100%"><defs>{defs}</defs><rect width="100%" height="100%" fill="{paint}"/></svg>'
            return '',svg
        return 'background:'+(paint if paint!='none' else '#FFFFFF')+';', ''

    def convert(self):
        output=[]
        slide_records=[]
        prels=self.rels('ppt/presentation.xml')
        for idx,item in enumerate(allof(self.presentation,'p:sldIdLst/p:sldId'),1):
            self.slide_no=idx; self.oid=f's{idx:03d}'
            target=prels[item.get('{'+NS['r']+'}id')]['target']
            root=self.xml(target)
            layoutpart=self.related(target,'slideLayout')
            masterpart=self.related(layoutpart,'slideMaster')
            self.layout,self.master=self.xml(layoutpart),self.xml(masterpart)
            self.setup_theme(masterpart,root,self.layout)
            before=dict(self.stats); textstart=len(self.texts)
            if find(root,'p:timing') is not None: self.issue('animations','PPTX animations are not converted','warning')
            if find(root,'p:transition') is not None: self.issue('transition','PPTX transitions are not converted','warning')
            if root.get('show','1')=='0': self.issue('hidden-slide','Hidden PPTX slide is included to preserve slide count','warning')
            bgstyle,bgsvg=self.background([root,self.layout,self.master])
            layers=[]
            # Template artwork is painted first. Layout/master placeholders are style sources only.
            if root.get('showMasterSp','1') != '0':
                if attr(self.layout,'showMasterSp','1') != '0': layers.append(('master',masterpart,self.master))
                layers.append(('layout',layoutpart,self.layout))
            layers.append(('slide',target,root))
            content=bgsvg
            for origin,part,layer in layers:
                tree=find(layer,'p:cSld/p:spTree')
                for shape in tree if tree is not None else []:
                    kind=tag(shape)
                    if kind in ('nvGrpSpPr','grpSpPr','extLst'): continue
                    if origin!='slide' and self.ph(shape) is not None: continue
                    content+=self.render_shape(shape,origin,part)
            title=attr(find(root,'p:cSld'),'name','') or '슬라이드 '+str(idx)
            output.append(f'<section class="ppt-slide" id="slide-{idx}" data-slide-index="{idx}" aria-label="{esc(title)}" style="width:{nfmt(self.width)}px;height:{nfmt(self.height)}px;{bgstyle}">{content}</section>')
            notes=''
            if self.include_notes:
                notesroot=self.xml(self.related(target,'notesSlide'))
                noteparas=[]
                for sp in allof(notesroot,'p:cSld/p:spTree/p:sp'):
                    if attr(self.ph(sp),'type')=='body':
                        noteparas += [''.join(t.text or '' for t in allof(p,'.//a:t')) for p in allof(sp,'p:txBody/a:p')]
                notes='\n'.join(noteparas)
            self.notes.append(notes)
            slide_records.append({'index':idx,'source_part':target,'objects':self.stats['objects']-before.get('objects',0),'text_runs':len(self.texts)-textstart})
        self.oid='package'
        for unused in sorted(set(self.fallbacks)-self.used_fallbacks): self.issue('fallback-manifest','Unknown or unused fallback object_id: '+unused)
        if self.stats['outer_shadows']: self.issue('shadow-rendering','Outer shadows use CSS filter approximation; compare with PowerPoint','warning')
        if any(n.startswith('ppt/embeddings/') for n in self.z.namelist()): self.issue('embedded-files','Embedded files are not attached to output','warning')
        if any('comments/' in n for n in self.z.namelist()): self.issue('comments','Review comments are excluded','warning')
        manifest={'title':self.source.stem,'width':self.width,'height':self.height,'slide_count':len(output),'fonts':sorted(self.fonts),'issues':self.issues,'notes':self.notes if self.include_notes else [],'source_sha256':hashlib.sha256(self.source.read_bytes()).hexdigest()}
        report={'schema_version':1,'source':self.source.name,'source_sha256':manifest['source_sha256'],'slide_count':len(output),'width_px':self.width,'height_px':self.height,'stats':dict(self.stats),'fonts_requested':sorted(self.fonts),'notes_included':self.include_notes,'issues':self.issues,'slides':slide_records,'text_runs':len(self.texts),'text_sha256':hashlib.sha256('\0'.join(self.texts).encode()).hexdigest(),'powerpoint_visual_comparison':'NOT_PERFORMED','browser_layout_check':'NOT_PERFORMED'}
        return ''.join(output),manifest,report
