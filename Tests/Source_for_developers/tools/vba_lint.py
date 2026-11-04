import re,glob,sys
files=sorted(glob.glob("/data/v82/src/VBA_Operator/*.bas"))+sorted(glob.glob("/data/v82/src/VBA_Viewer/*.bas"))+sorted(glob.glob("/data/v82/src/Object_Code/*.txt"))
issues=[];procs={}
def strip(l):
    # remove strings and comments
    out="";ins=False
    for ch in l:
        if ch=='"': ins=not ins; out+='"'; continue
        if ins: continue
        if ch=="'": break
        out+=ch
    return out
for f in files:
    raw=open(f,encoding="utf-8").read()
    try: raw.encode("cp1251")
    except UnicodeEncodeError as e: issues.append(f"{f}: не CP1251 символ {raw[e.start]!r} pos {e.start}")
    if "\ufffd" in raw: issues.append(f"{f}: символ замены U+FFFD")
    lines=raw.split("\n"); stack=[]; cont=0; logical=""
    for i,l in enumerate(lines,1):
        if len(l)>1000: issues.append(f"{f}:{i}: строка длиннее 1000")
        s=strip(l).strip()
        if s.endswith(" _"):
            cont+=1; logical+=s[:-1]+" "; 
            if cont>24: issues.append(f"{f}:{i}: >24 переносов")
            continue
        s=logical+s; logical=""; cont=0
        if not s: continue
        m=re.match(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property (Get|Let|Set)) (\w+)',s)
        if m:
            kind=m.group(3).split()[0]; stack.append((kind,i)); 
            procs[m.group(5).lower()]=(f,m.group(1) or "Public ")
            continue
        if re.match(r'^End (Sub|Function|Property)\b',s):
            k=s.split()[1]
            if not stack or stack[-1][0]!=k: issues.append(f"{f}:{i}: End {k} без пары, стек {stack[-3:]}")
            else: stack.pop()
            if stack: issues.append(f"{f}:{i}: незакрытые блоки {stack}"); stack=[]
            continue
        # statements split by ':' (rough)
        if re.match(r'^If\b.*\bThen$',s): stack.append(("If",i)); continue
        if re.match(r'^(ElseIf\b.*Then|Else)$',s):
            if not stack or stack[-1][0]!="If": issues.append(f"{f}:{i}: Else вне If")
            continue
        if s=="End If":
            if not stack or stack[-1][0]!="If": issues.append(f"{f}:{i}: End If без If {stack[-1:] }")
            else: stack.pop()
            continue
        for kw,end in (("For","Next"),("Do","Loop"),("With","End With"),("Select Case","End Select"),("While","Wend")):
            pass
        if re.match(r'^For\b',s) and not re.search(r':\s*Next\b',s): stack.append(("For",i))
        elif re.match(r'^Next\b',s):
            if not stack or stack[-1][0]!="For": issues.append(f"{f}:{i}: Next без For {stack[-1:]}")
            else: stack.pop()
        elif re.match(r'^Do\b',s): stack.append(("Do",i))
        elif re.match(r'^Loop\b',s):
            if not stack or stack[-1][0]!="Do": issues.append(f"{f}:{i}: Loop без Do")
            else: stack.pop()
        elif re.match(r'^With\b',s): stack.append(("With",i))
        elif s=="End With":
            if not stack or stack[-1][0]!="With": issues.append(f"{f}:{i}: End With без With")
            else: stack.pop()
        elif re.match(r'^Select Case\b',s): stack.append(("Select",i))
        elif s=="End Select":
            if not stack or stack[-1][0]!="Select": issues.append(f"{f}:{i}: End Select без Select")
            else: stack.pop()
        if re.match(r'^\s*Dim ',l) and len(l)-len(l.lstrip())>=8: issues.append(f"{f}:{i}: Dim внутри блока")
        if re.search(r'\bThen ReDim\b',s): issues.append(f"{f}:{i}: однострочный If ... ReDim")
    if stack: issues.append(f"{f}: в конце незакрытые {stack}")
# OnAction targets
alltxt="\n".join(open(f,encoding="utf-8").read() for f in files)
targets=set(re.findall(r'(?:Btn|VBtn|AddMenuItem)\s+[^\n]*?"(\w+)"\s*,\s*\d+',alltxt))
targets|=set(re.findall(r'pre & "(\w+)"',alltxt))
for m in re.finditer(r'\n\s*(?:Btn|VBtn)\s+\w+,\s*"[^"]*",\s*"[^"]*",\s*"(\w+)"',alltxt): targets.add(m.group(1))
for tname in sorted(targets):
    p=procs.get(tname.lower())
    if not p: issues.append(f"OnAction: процедура {tname} не найдена")
    elif p[1].strip()=="Private": issues.append(f"OnAction: {tname} — Private")
print("files",len(files),"procs",len(procs),"onaction targets",len(targets))
print("\n".join(issues) if issues else "OK: проблем не найдено")
