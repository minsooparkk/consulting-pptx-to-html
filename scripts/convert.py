#!/usr/bin/env python3
"""Convert an approved PPTX without rewriting content or changing slide count."""
from __future__ import annotations
import argparse
import hashlib
import html
import json
import math
import re
import sys
from html.parser import HTMLParser
from pathlib import Path
from pptx_html import Converter, data_uri, css_family, safe_json
from procedural import build_procedural_reveal, positive_interval

ROOT = Path(__file__).resolve().parents[1]

class RunCollector(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.depth=0; self.current=[]; self.runs=[]; self.slides=0
    def handle_starttag(self, tag, attrs):
        a=dict(attrs)
        if tag=='section' and 'ppt-slide' in a.get('class','').split(): self.slides+=1
        if tag=='span':
            if self.depth: self.depth+=1
            elif 'ppt-run' in a.get('class','').split(): self.depth=1; self.current=[]
    def handle_endtag(self, tag):
        if tag=='span' and self.depth:
            self.depth-=1
            if self.depth==0: self.runs.append(''.join(self.current))
    def handle_data(self, text):
        if self.depth: self.current.append(text)

def main(argv=None):
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('pptx',type=Path)
    ap.add_argument('-o','--output',required=True,type=Path)
    ap.add_argument('--report',type=Path,help='Default: OUTPUT.audit.json')
    ap.add_argument('--include-notes',action='store_true',help='Notes become readable by every HTML recipient')
    ap.add_argument('--fallback-manifest',type=Path,help='JSON object ID -> local image path; relative to manifest')
    ap.add_argument('--font',action='append',default=[],metavar='FAMILY,WEIGHT,FILE',help='Embed a licensed font, e.g. Pretendard,400,/fonts/Pretendard-Regular.woff2')
    ap.add_argument('--motion',choices=['auto','none'],default='auto',help='Automatic brief entrances on slide navigation (default), or no added motion')
    ap.add_argument('--procedure-manifest',type=Path,help='JSON explicit ordered object groups for selected slides')
    ap.add_argument('--procedure-interval-ms',type=positive_interval,default=300,help='Positive interval between procedural groups in milliseconds (default: 300)')
    ap.add_argument('--line-height-factor',action='append',default=[],metavar='FAMILY=FACTOR',help='Measured natural line-height factor for percentage spacing; Pretendard defaults to 1.2')
    ap.add_argument('--allow-unsupported',action='store_true',help='Keep a marked partial preview and return 0 even with conversion errors')
    ap.add_argument('--overwrite',action='store_true')
    args=ap.parse_args(argv)
    if args.procedure_manifest and args.motion=='none': ap.error('--procedure-manifest cannot be combined with --motion none')
    if args.pptx.suffix.lower()!='.pptx': ap.error('Only .pptx is accepted; convert legacy .ppt in PowerPoint first')
    if args.output.suffix.lower() not in ('.html','.htm'): ap.error('Output must be .html or .htm')
    report_path=args.report or args.output.with_suffix('.audit.json')
    if args.output.resolve()==args.pptx.resolve() or report_path.resolve()==args.pptx.resolve() or report_path.resolve()==args.output.resolve(): ap.error('Input, HTML and report paths must differ')
    if args.procedure_manifest and args.procedure_manifest.resolve() in (args.output.resolve(),report_path.resolve()): ap.error('Procedure manifest cannot be overwritten by HTML or report output')
    for p in [args.output,report_path]:
        if p.exists() and not args.overwrite: ap.error(f'Already exists: {p.name}. Use --overwrite intentionally')
    fallbacks={}
    if args.fallback_manifest:
        entries=json.loads(args.fallback_manifest.read_text(encoding='utf-8'))
        if not isinstance(entries,dict): ap.error('Fallback manifest must be a JSON object')
        fallbacks={k:str((args.fallback_manifest.parent/Path(v)).resolve()) for k,v in entries.items()}
    fontcss=[]; embedded=[]
    for spec in args.font:
        parts=spec.split(',',2)
        if len(parts)!=3 or not re.fullmatch(r'[1-9]\d{0,2}',parts[1]): ap.error('--font expects FAMILY,WEIGHT,FILE')
        family,weight,filename=parts; file=Path(filename)
        mime={'.woff2':'font/woff2','.woff':'font/woff','.ttf':'font/ttf','.otf':'font/otf'}.get(file.suffix.lower())
        if not mime: ap.error('Supported fonts: woff2, woff, ttf, otf')
        fontcss.append('@font-face{font-family:'+css_family(family)+';font-style:normal;font-weight:'+weight+';src:url("'+data_uri(file.read_bytes(),mime)+'");font-display:block;}')
        embedded.append({'family':family,'weight':weight})
    factors={}
    for spec in args.line_height_factor:
        try:
            family,value=spec.rsplit('=',1)
            factor=float(value)
            if not family.strip() or not math.isfinite(factor) or factor<=0: raise ValueError()
        except ValueError: ap.error('--line-height-factor expects FAMILY=FACTOR with a finite positive number')
        factors[family.strip()]=factor
    converter=Converter(args.pptx,args.include_notes,fallbacks,line_height_factors=factors)
    try: slides,manifest,report=converter.convert()
    finally: converter.close()
    collector=RunCollector(); collector.feed(slides)
    digest=hashlib.sha256('\0'.join(collector.runs).encode()).hexdigest()
    content_ok=digest==report['text_sha256'] and collector.slides==report['slide_count']
    if not content_ok: raise RuntimeError('HTML serialization text/slide mismatch; output aborted')
    report['html_text_sha256']=digest
    report['serialization_check']='PASS'
    report['embedded_fonts']=embedded
    report['font_availability']='EMBEDDED_SUBSET' if embedded else 'SYSTEM_FONTS_NOT_VERIFIED'
    report['source_bytes']=args.pptx.stat().st_size
    errors=[x for x in report['issues'] if x['severity']=='error']
    report['conversion_status']='PARTIAL_REVIEW_REQUIRED' if errors else 'PASS_STRUCTURE'
    manifest['conversion_status']=report['conversion_status']
    manifest['embedded_fonts']=embedded
    manifest['motion']={'mode':args.motion,'duration_ms':420,'max_delay_ms':500}
    report['motion']={**manifest['motion'],'maximum_total_ms':920 if args.motion=='auto' else 0,'manual_steps':False,'item_click_actions':False,'browser_verified':False}
    if args.procedure_manifest:
        procedure_data=json.loads(args.procedure_manifest.read_text(encoding='utf-8'))
        procedure,procedure_audit=build_procedural_reveal(procedure_data,slides,args.procedure_interval_ms)
        manifest['procedural_reveal']=procedure
        report['procedural_reveal']=procedure_audit
        report['motion']['normal_maximum_total_ms']=report['motion']['maximum_total_ms']
        report['motion']['maximum_total_ms']=max(report['motion']['maximum_total_ms'],procedure_audit['maximum_total_ms'])
    template=(ROOT/'assets/player.html').read_text(encoding='utf-8')
    replacements={'TITLE':html.escape(manifest['title'],quote=True),'STYLE':'\n'.join(fontcss)+'\n'+(ROOT/'assets/player.css').read_text(encoding='utf-8'),
                  'SLIDES':slides,'MANIFEST':safe_json(manifest),'SCRIPT':(ROOT/'assets/player.js').read_text(encoding='utf-8')}
    if args.motion=='auto':
        replacements['STYLE']+='\n'+(ROOT/'assets/auto-motion.css').read_text(encoding='utf-8')
        replacements['SCRIPT']+='\n'+(ROOT/'assets/auto-motion.js').read_text(encoding='utf-8')
    # One-pass token replacement: source text containing {{...}} cannot become executable CSS/JS.
    out=re.sub(r'\{\{(TITLE|STYLE|SLIDES|MANIFEST|SCRIPT)\}\}',lambda m:replacements[m.group(1)],template)
    if re.search(r'\{\{(TITLE|STYLE|SLIDES|MANIFEST|SCRIPT)\}\}',template) is None: raise RuntimeError('Invalid player template')
    report['html_bytes']=len(out.encode('utf-8'))
    report['html_sha256']=hashlib.sha256(out.encode('utf-8')).hexdigest()
    args.output.parent.mkdir(parents=True,exist_ok=True); report_path.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_bytes(out.encode('utf-8'))
    report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'html':str(args.output),'report':str(report_path),'slides':report['slide_count'],'objects':report['stats'].get('objects',0),'text_runs':report['text_runs'],'errors':len(errors),'warnings':len(report['issues'])-len(errors),'status':report['conversion_status'],'visual_comparison':'NOT_PERFORMED'},ensure_ascii=False))
    return 2 if errors and not args.allow_unsupported else 0

if __name__=='__main__':
    if hasattr(sys.stdout,'reconfigure'): sys.stdout.reconfigure(encoding='utf-8')
    try: sys.exit(main())
    except (ValueError,OSError,KeyError) as exc:
        print('Conversion failed: '+str(exc),file=sys.stderr); sys.exit(1)
