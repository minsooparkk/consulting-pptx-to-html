#!/usr/bin/env python3
"""Measure generated HTML in a local headless Chromium browser, without installing packages."""
from __future__ import annotations
import argparse
import html
import json
import os
import re
import shutil
import subprocess
import sys
import uuid
from pathlib import Path

PROBE='''<script>
window.PPTPlayer.ready.then(() => window.PPTPlayer.audit()).then(result => {
 const node=document.createElement('pre');node.id='ppt-browser-audit';node.dataset.auditResult='true';
 node.textContent=JSON.stringify(result);document.body.append(node);
}).catch(error => {
 const node=document.createElement('pre');node.id='ppt-browser-audit';node.dataset.auditResult='true';
 node.textContent=JSON.stringify({ok:false,error:String(error)});document.body.append(node);
});
</script>'''

def browser_path(explicit=None):
    if explicit: return str(explicit.expanduser().resolve())
    candidates=[]
    for name in ['google-chrome','chromium','chromium-browser','msedge','chrome']:
        found=shutil.which(name)
        if found: candidates.append(found)
    candidates += ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
                   '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
                   '/Applications/Chromium.app/Contents/MacOS/Chromium']
    for env in ['PROGRAMFILES','PROGRAMFILES(X86)','LOCALAPPDATA']:
        if os.environ.get(env):
            candidates += [str(Path(os.environ[env])/'Google/Chrome/Application/chrome.exe'),
                           str(Path(os.environ[env])/'Microsoft/Edge/Application/msedge.exe')]
    return next((str(p) for p in candidates if Path(p).is_file()),None)

def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('html',type=Path)
    parser.add_argument('--report',type=Path)
    parser.add_argument('--browser',type=Path,help='Chrome/Chromium/Edge executable')
    parser.add_argument('--timeout',type=int,default=60)
    args=parser.parse_args(argv)
    file=args.html.resolve()
    src=file.read_text(encoding='utf-8')
    if 'id="ppt-manifest"' not in src or 'PPTPlayer' not in src: parser.error('Expected this skill\'s generated HTML')
    browser=browser_path(args.browser)
    if not browser: parser.error('Chrome/Chromium/Edge not found. Provide --browser, or run await window.PPTPlayer.audit() in the browser console')
    report=args.report or file.with_suffix('.browser-qa.json')
    if report.resolve()==file: parser.error('Report cannot overwrite the HTML')
    work=file.parent/('.pptx-html-audit-'+uuid.uuid4().hex)
    # Inherit host ACLs rather than tempfile's restrictive Windows mode=0700.
    work.mkdir()
    try:
        probe=work/'probe.html'
        probe.write_text(re.sub(r'</body>',lambda _:PROBE+'</body>',src,count=1,flags=re.I),encoding='utf-8')
        command=[browser,'--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check',
                 '--window-size=1440,900','--virtual-time-budget=15000','--dump-dom',
                 '--user-data-dir='+str(work/'profile'),probe.as_uri()]
        proc=subprocess.run(command,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=args.timeout)
        match=re.search(r'<pre[^>]*id="ppt-browser-audit"[^>]*>(.*?)</pre>',proc.stdout,re.S)
        if not match:
            raise RuntimeError('Browser did not return audit data. Check browser execution permissions or use the in-page audit API. Exit='+str(proc.returncode))
        result=json.loads(html.unescape(match.group(1)))
        result['checked_file']=file.name
        report.parent.mkdir(parents=True,exist_ok=True)
        report.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        print(json.dumps({'ok':result.get('ok',False),'slides':len(result.get('slides',[])),'report':str(report),'powerpoint_visual_comparison':'NOT_PERFORMED'},ensure_ascii=False))
        return 0 if result.get('ok') else 2
    finally:
        shutil.rmtree(work,ignore_errors=True)

if __name__=='__main__':
    if hasattr(sys.stdout,'reconfigure'): sys.stdout.reconfigure(encoding='utf-8')
    try: sys.exit(main())
    except (OSError,ValueError,RuntimeError,subprocess.TimeoutExpired) as exc:
        print('Browser audit failed: '+str(exc),file=sys.stderr);sys.exit(1)
