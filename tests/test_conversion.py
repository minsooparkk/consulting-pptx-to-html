"""Run with: python -m unittest discover -s tests -v (stdlib only)."""
import json
import hashlib
import subprocess
import sys
import shutil
import uuid
from contextlib import contextmanager
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from pptx_html import Converter, NS
from convert import RunCollector

SAMPLE=ROOT/'examples/sample.pptx'

@contextmanager
def workspace():
    # Inherit ACLs; Python 3.13 TemporaryDirectory(mode=0700) can lock out sandboxed Windows tokens.
    folder=ROOT/'tests'/('.tmp-'+uuid.uuid4().hex)
    folder.mkdir()
    try: yield str(folder)
    finally: shutil.rmtree(folder)

class ConversionTests(unittest.TestCase):
    def parse(self, path=SAMPLE, **kw):
        c=Converter(path,**kw)
        try:
            out,manifest,report=c.convert()
            return out,manifest,report,list(c.texts)
        finally: c.close()

    def mutated(self, folder, fragment, rel=None, binary=None):
        c=Converter(SAMPLE)
        try:
            rid=c.presentation.find('p:sldIdLst/p:sldId',NS).get('{'+NS['r']+'}id')
            part=c.rels('ppt/presentation.xml')[rid]['target']
        finally: c.close()
        file=Path(folder)/'fixture.pptx'
        with ZipFile(SAMPLE) as src,ZipFile(file,'w') as dst:
            for item in src.infolist():
                data=src.read(item.filename)
                if item.filename==part:
                    node=ET.fromstring(data); tree=node.find('p:cSld/p:spTree',NS)
                    wrapper=ET.fromstring('<wrapper xmlns:p="'+NS['p']+'" xmlns:a="'+NS['a']+'" xmlns:r="'+NS['r']+'">'+fragment+'</wrapper>')
                    for child in wrapper: tree.append(child)
                    data=ET.tostring(node,encoding='utf-8',xml_declaration=True)
                if rel and item.filename==str(Path(part).parent/'_rels'/(Path(part).name+'.rels')).replace('\\','/'):
                    node=ET.fromstring(data); node.append(ET.fromstring(rel)); data=ET.tostring(node)
                dst.writestr(item,data)
            if binary:
                for name,data in binary.items():dst.writestr(name,data)
        return file

    def test_sample_structure_and_text(self):
        out,m,r,texts=self.parse()
        parsed=RunCollector(); parsed.feed(out)
        self.assertEqual(r['slide_count'],4)
        self.assertEqual(parsed.slides,4)
        self.assertEqual(parsed.runs,texts)
        self.assertFalse([i for i in r['issues'] if i['severity']=='error'])
        self.assertEqual(r['fonts_requested'],['Pretendard'])
        self.assertEqual(r['stats']['tables'],1)
        self.assertEqual(r['stats']['gradients'],4)
        self.assertIn('data-object-id="s002_layout_2"',out)
        self.assertIn('data-object-id="s002_master_2"',out)
        self.assertEqual(r['powerpoint_visual_comparison'],'NOT_PERFORMED')

    def test_independent_visible_source_text_inventory(self):
        c=Converter(SAMPLE); expected=[]
        try:
            for item in c.presentation.findall('p:sldIdLst/p:sldId',NS):
                part=c.rels('ppt/presentation.xml')[item.get('{'+NS['r']+'}id')]['target']
                lp=c.related(part,'slideLayout'); mp=c.related(lp,'slideMaster')
                for origin,p in [('master',mp),('layout',lp),('slide',part)]:
                    root=c.xml(p)
                    for shape in root.findall('p:cSld/p:spTree/*',NS):
                        if origin!='slide' and shape.find('.//p:nvPr/p:ph',NS) is not None: continue
                        expected.extend(t.text or '' for t in shape.findall('.//a:t',NS))
            out,_,_=c.convert(); parsed=RunCollector();parsed.feed(out)
            self.assertEqual(parsed.runs,expected)
        finally:c.close()

    def test_notes_are_opt_in(self):
        _,m,r,_=self.parse()
        self.assertEqual(m['notes'],[])
        self.assertFalse(r['notes_included'])
        _,m,r,_=self.parse(include_notes=True)
        self.assertEqual(len(m['notes']),4)
        self.assertTrue(r['notes_included'])

    def test_escape_and_group_coordinate_mapping(self):
        fragment='''<p:grpSp><p:nvGrpSpPr><p:cNvPr id="101" name="group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="102" name="text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="500000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>&lt;script&gt;unsafe&lt;/script&gt; &amp; {{SCRIPT}}</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>'''
        with workspace() as tmp:
            path=self.mutated(tmp,fragment)
            out,m,r,texts=self.parse(path)
            self.assertIn('&lt;script&gt;unsafe&lt;/script&gt; &amp; {{SCRIPT}}',out)
            self.assertNotIn('<script>unsafe',out)
            self.assertIn('left:20px;top:10px;',out)
            self.assertEqual(r['stats']['groups'],1)

    def test_unsupported_graphic_is_visible_and_reported(self):
        fragment='''<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="103" name="chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></p:graphicFrame>'''
        with workspace() as tmp:
            path=self.mutated(tmp,fragment)
            out,_,r,_=self.parse(path)
            self.assertIn('검토 필요: chart',out)
            self.assertTrue(any(i['feature']=='chart' and i['severity']=='error' for i in r['issues']))
            html=Path(tmp)/'result.html'
            result=subprocess.run([sys.executable,str(ROOT/'scripts/convert.py'),str(path),'-o',str(html)],capture_output=True,text=True,encoding='utf-8')
            self.assertEqual(result.returncode,2,result.stderr)
            self.assertEqual(json.loads(html.with_suffix('.audit.json').read_text(encoding='utf-8'))['conversion_status'],'PARTIAL_REVIEW_REQUIRED')

    def test_embedded_image_and_element_fallback(self):
        import base64
        png=base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aT6sAAAAASUVORK5CYII=')
        fragment='''<p:pic><p:nvPicPr><p:cNvPr id="104" name="embedded"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rImageTest"/><a:srcRect l="10000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>'''
        rel='<Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships" Id="rImageTest" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/test.png"/>'
        with workspace() as tmp:
            path=self.mutated(tmp,fragment,rel,{'ppt/media/test.png':png})
            out,_,r,_=self.parse(path)
            self.assertIn('data:image/png;base64,',out)
            self.assertIn('left:-11.1111%',out)
            self.assertEqual(r['stats']['images'],1)
            fallback=Path(tmp)/'fallback.png';fallback.write_bytes(png)
            out,_,r,_=self.parse(path,fallbacks={'s001_slide_104':str(fallback)})
            self.assertEqual(r['stats']['fallbacks'],1)
            self.assertTrue(any(i['feature']=='element-fallback' for i in r['issues']))

    def test_cli_one_pass_tokens_and_overwrite_guard(self):
        with workspace() as tmp:
            html=Path(tmp)/'sample.html'
            cmd=[sys.executable,str(ROOT/'scripts/convert.py'),str(SAMPLE),'-o',str(html)]
            p=subprocess.run(cmd,capture_output=True,text=True,encoding='utf-8')
            self.assertEqual(p.returncode,0,p.stderr)
            page=html.read_text(encoding='utf-8')
            self.assertIn('id="ppt-manifest"',page)
            self.assertNotIn('cdn.jsdelivr',page)
            r=json.loads(html.with_suffix('.audit.json').read_text(encoding='utf-8'))
            self.assertEqual(r['serialization_check'],'PASS')
            self.assertEqual(r['html_text_sha256'],r['text_sha256'])
            p=subprocess.run(cmd,capture_output=True,text=True,encoding='utf-8')
            self.assertNotEqual(p.returncode,0)

    def test_font_natural_line_height_and_exact_point_spacing(self):
        def shape(identifier, family, spacing):
            return f'''<p:sp><p:nvSpPr><p:cNvPr id="{identifier}" name="metric-fixture"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3800000" cy="1900000"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:lnSpc>{spacing}</a:lnSpc></a:pPr><a:r><a:rPr sz="1500"><a:latin typeface="{family}"/><a:ea typeface="{family}"/></a:rPr><a:t>First line</a:t></a:r><a:br/><a:r><a:rPr sz="1500"><a:latin typeface="{family}"/><a:ea typeface="{family}"/></a:rPr><a:t>Second line</a:t></a:r></a:p></p:txBody></p:sp>'''
        fragment=''.join([
            shape(110,'Pretendard','<a:spcPct val="110000"/>'),
            shape(111,'Pretendard','<a:spcPts val="1800"/>'),
            shape(112,'Arial','<a:spcPct val="110000"/>')])
        def paragraph(page, identifier):return page.split(f'id="s001_slide_{identifier}"',1)[1].split('</p>',1)[0]
        with workspace() as tmp:
            fixture=self.mutated(tmp,fragment)
            out,_,report,_=self.parse(fixture)
            self.assertIn('line-height:1.32;',paragraph(out,110))
            self.assertIn('line-height:24px;',paragraph(out,111))
            self.assertIn('line-height:1.1;',paragraph(out,112))
            out,_,report,_=self.parse(fixture,line_height_factors={'Pretendard':1.0,'Arial':1.15})
            self.assertIn('line-height:1.1;',paragraph(out,110))
            self.assertIn('line-height:24px;',paragraph(out,111))
            self.assertIn('line-height:1.265;',paragraph(out,112))
            self.assertTrue(report['line_height_calibration']['exact_point_spacing_unchanged'])

    def test_auto_motion_opt_out_preserves_content_and_file_hash(self):
        with workspace() as tmp:
            runs=[]
            for mode in ['auto','none']:
                target=Path(tmp)/(mode+'.html')
                result=subprocess.run([sys.executable,str(ROOT/'scripts/convert.py'),str(SAMPLE),'-o',str(target),'--motion',mode],capture_output=True,text=True,encoding='utf-8')
                self.assertEqual(result.returncode,0,result.stderr)
                report=json.loads(target.with_suffix('.audit.json').read_text(encoding='utf-8'))
                page=target.read_text(encoding='utf-8'); parser=RunCollector();parser.feed(page);runs.append(parser.runs)
                self.assertEqual(report['motion']['mode'],mode)
                self.assertFalse(report['motion']['manual_steps'])
                self.assertEqual(report['html_sha256'],hashlib.sha256(target.read_bytes()).hexdigest())
                self.assertEqual(report['html_bytes'],target.stat().st_size)
                self.assertEqual('content-arrive' in page,mode=='auto')
            self.assertEqual(runs[0],runs[1])

    def test_invalid_line_height_override_rejected_without_output(self):
        with workspace() as tmp:
            for value in ['Pretendard=nan','Pretendard=0','=1.2','Pretendard=inf']:
                target=Path(tmp)/'invalid.html'
                result=subprocess.run([sys.executable,str(ROOT/'scripts/convert.py'),str(SAMPLE),'-o',str(target),'--line-height-factor',value],capture_output=True,text=True,encoding='utf-8')
                self.assertNotEqual(result.returncode,0)
                self.assertFalse(target.exists())

if __name__=='__main__': unittest.main()
