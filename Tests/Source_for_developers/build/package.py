import os, shutil, glob, openpyxl, datetime
from openpyxl.styles import Font
V="/data/v82"; P=V+"/package/ELKA_v8_2"; shutil.rmtree(V+"/package",ignore_errors=True); os.makedirs(P)
def crlf(t): return t.replace("\r\n","\n").replace("\n","\r\n")
def wtext(path,text,enc="utf-8",bom=False):
    os.makedirs(os.path.dirname(path),exist_ok=True)
    data=crlf(text).encode(enc)
    if bom: data=b"\xef\xbb\xbf"+data
    open(path,"wb").write(data)
# cmd
for f in glob.glob(V+"/pkg_src/*.cmd"):
    wtext(P+"/"+os.path.basename(f),open(f,encoding="utf-8").read(),"ascii")
# ps1 with BOM
for f in glob.glob(V+"/pkg_src/Setup/*.ps1"):
    wtext(P+"/Setup/"+os.path.basename(f),open(f,encoding="utf-8-sig").read(),bom=True)
# workbooks
os.makedirs(P+"/Setup/Workbooks")
for f in ["ELKA_Operator_Template.xlsx","ELKA_Operator_Demo.xlsx","ELKA_Viewer_Template.xlsx","ELKA_Viewer_Demo.xlsx"]:
    shutil.copy(V+"/out/"+f,P+"/Setup/Workbooks/"+f)
# VBA utf-8 + CP1251 copies
for sub in ["VBA_Operator","VBA_Viewer"]:
    for f in sorted(glob.glob(V+"/src/"+sub+"/*.bas")):
        t=open(f,encoding="utf-8").read()
        wtext(P+"/Setup/"+sub+"/"+os.path.basename(f),t)
        wtext(P+"/Setup/Manual_Import_CP1251/"+sub+"/"+os.path.basename(f),t,"cp1251")
for f in glob.glob(V+"/src/Object_Code/*.txt"):
    t=open(f,encoding="utf-8").read()
    wtext(P+"/Setup/Object_Code/"+os.path.basename(f),t)
    wtext(P+"/Setup/Manual_Import_CP1251/"+os.path.basename(f),t,"cp1251")
# demo
shutil.copytree(V+"/lotest_arch/Архив",P+"/Demo/Архив")
os.makedirs(P+"/Demo/Publish"); shutil.copy(V+"/out/ELKA_data_v000001.xlsx",P+"/Demo/Publish/")
# sample 1C export
os.makedirs(P+"/Tests")
shutil.copy(V+"/smoke_1c.xlsx",P+"/Demo/Sample_1C_export.xlsx")
shutil.copy(V+"/smoke_1c.xlsx",P+"/Tests/Sample_1C_export.xlsx")
# docs
os.makedirs(P+"/Docs")
for f in glob.glob(V+"/docs/out/*"): shutil.copy(f,P+"/Docs/")
# tests
names={"full_run_limit.txt":"00_Full_test_not_completed.txt", "runtime_report.txt":"01_Repair_smoke_test.txt", "runtime_messages.txt":"02_Runtime_messages.txt", "Viewer_test_log.txt":"03_Viewer_test_log.txt", "vba_lint.txt":"04_VBA_static_check.txt", "independent_ledger.json":"05_Independent_ledger.json", "release_summary.txt":"06_Release_summary.txt", "1c_import.log":"07_1C_import.txt", "workbook_audit_summary.txt":"08_Workbook_audit.txt"}
for a,b in names.items(): wtext(P+"/Tests/"+b,open(V+"/tests_out/"+a,encoding="utf-8").read(),bom=True)
for d in ["harness","tools","build"]:
    shutil.copytree(V+"/"+d,P+"/Tests/Source_for_developers/"+d,ignore=shutil.ignore_patterns("__pycache__","*.pyc","data_demo.json"))
wtext(P+"/README_RU.txt",open(V+"/pkg_src/README_RU.txt",encoding="utf-8").read(),bom=True)
print("ok")
# viewer-only package
VP=V+"/package/ELKA_v8_2_Viewer"; os.makedirs(VP+"/Setup/Workbooks"); os.makedirs(VP+"/Docs")
shutil.copy(P+"/3_INSTALL_VIEWER.cmd",VP)
for f in ["Common.ps1","Install_Viewer.ps1"]: shutil.copy(P+"/Setup/"+f,VP+"/Setup/")
for f in ["ELKA_Viewer_Template.xlsx","ELKA_Viewer_Demo.xlsx"]: shutil.copy(P+"/Setup/Workbooks/"+f,VP+"/Setup/Workbooks/")
shutil.copytree(P+"/Setup/VBA_Viewer",VP+"/Setup/VBA_Viewer")
os.makedirs(VP+"/Setup/Object_Code"); shutil.copy(P+"/Setup/Object_Code/Viewer_ThisWorkbook.txt",VP+"/Setup/Object_Code/")
os.makedirs(VP+"/Setup/Manual_Import_CP1251"); shutil.copytree(P+"/Setup/Manual_Import_CP1251/VBA_Viewer",VP+"/Setup/Manual_Import_CP1251/VBA_Viewer")
shutil.copy(P+"/Setup/Manual_Import_CP1251/Viewer_ThisWorkbook.txt",VP+"/Setup/Manual_Import_CP1251/")
for f in ["01_Install_RU.pdf","03_User_RU.pdf"]: shutil.copy(P+"/Docs/"+f,VP+"/Docs/")
wtext(VP+"/README_RU.txt",open(V+"/pkg_src/README_VIEWER_RU.txt",encoding="utf-8").read(),bom=True)
print("viewer ok")
