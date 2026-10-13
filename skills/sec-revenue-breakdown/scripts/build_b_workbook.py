"""Single repeatable B CLI: plan -> selected template -> text patch -> strict QA.

Existing files are never overwritten. Runtime dependencies are reused, not
installed. Exit status from authoring is retained in evidence; a nonzero author
exit is accepted only if a complete expected XLSX passes independent strict QA.
Native Excel numeric verification is a separate final acceptance step.
"""
from pathlib import Path
import argparse,hashlib,json,os,shutil,subprocess,sys,tempfile
from b_workbook import make_plan
from preserve_excel_text import preserve,install_formulas
from validate_workbook_format import validate_with_plan

ROOT=Path(__file__).resolve().parents[1]

def run(args):
    out=args.out.resolve()
    if out.exists():raise FileExistsError(out)
    load=lambda p:json.loads(p.read_text(encoding='utf-8'))
    catalog=load(ROOT/'assets/templates/catalog.json')
    entry=next((t for t in catalog['templates'] if t['id']==args.template_id),None)
    if not entry:raise ValueError('Template ID not in catalog')
    template=ROOT/'assets/templates'/entry['file']
    if hashlib.sha256(template.read_bytes()).hexdigest()!=entry['sha256']:raise ValueError('Template changed; verify and update catalog before using')
    data,views,manifest=map(load,(args.breakdown,args.views,args.manifest))
    integrity=[]
    if args.verify_sources:
        for s in manifest['sources']:
            p=(args.manifest.parent/s['filename']).resolve()
            if not p.is_relative_to(args.manifest.parent.resolve()):raise ValueError('Source path outside package')
            if not p.is_file() or hashlib.sha256(p.read_bytes()).hexdigest()!=s['sha256']:raise ValueError('Source missing/changed: '+s['source_id'])
            integrity.append(s['source_id'])
    plan=make_plan(data,views,manifest,args.template_id)
    plan['input_sha256']={str(f.resolve()):hashlib.sha256(f.read_bytes()).hexdigest() for f in (args.breakdown,args.views,args.manifest)}
    plan['source_integrity']=dict(verified_count=len(integrity),checked=bool(args.verify_sources))
    out.parent.mkdir(parents=True,exist_ok=True);plan_file=out.with_suffix('.plan.json')
    if plan_file.exists():raise FileExistsError(plan_file)
    plan_file.write_text(json.dumps(plan,ensure_ascii=False,indent=2),encoding='utf-8')
    # Separate temporary runtime avoids package installation in shared skill.
    runtime_root=ROOT/'.runtime';runtime_root.mkdir(exist_ok=True)
    runtime=Path(tempfile.mkdtemp(prefix='b-author-',dir=runtime_root));link=runtime/'node_modules'
    try:
        deps=args.node_modules.resolve()
        if not (deps/'@oai/artifact-tool').is_dir():raise FileNotFoundError('artifact-tool runtime unavailable')
        if os.name=='nt':
            subprocess.run(['cmd','/c','mklink','/J',str(link),str(deps)],check=True,capture_output=True)
        else:link.symlink_to(deps,target_is_directory=True)
        builder=runtime/'build_b_workbook.mjs';shutil.copy2(ROOT/'scripts/build_b_workbook.mjs',builder)
        result=subprocess.run(['node',str(builder),'--plan',str(plan_file),'--template',str(template),'--out',str(out),'--defer-calculation'],capture_output=True,text=True,encoding='utf-8',errors='replace')
        if not out.is_file():raise RuntimeError('Authoring failed: '+result.stderr[-2000:])
        installed=install_formulas(out,plan)
        preserved=preserve(out)
        report=validate_with_plan(out,plan)
        report.update(template_id=args.template_id,template_sha256=entry['sha256'],author_exit_code=result.returncode,
                      formulas_installed=installed,whitespace_nodes_preserved=preserved,source_integrity=plan['source_integrity'],input_sha256=plan['input_sha256'],
                      acceptance='structural only; native Excel calculation still required')
        if result.returncode:report['author_warning']='Author process returned nonzero after export; independent strict validation passed. Native acceptance required.'
        out.with_suffix('.build.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(report,ensure_ascii=True))
    finally:
        # Remove only this task-created junction, then owned temporary files.
        if not runtime.resolve().is_relative_to(runtime_root.resolve()):
            raise RuntimeError('Refusing cleanup outside the owned runtime directory')
        if link.exists():
            if os.name=='nt':os.rmdir(link)
            else:link.unlink()
        shutil.rmtree(runtime)

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    for n in ['breakdown','views','manifest','out']:p.add_argument('--'+n,required=True,type=Path)
    p.add_argument('--template-id',choices=['flat','hierarchy','segment-op'],required=True)
    p.add_argument('--verify-sources',action='store_true',help='Verify every local source file SHA-256, without network')
    p.add_argument('--node-modules',type=Path,default=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules')
    run(p.parse_args())
