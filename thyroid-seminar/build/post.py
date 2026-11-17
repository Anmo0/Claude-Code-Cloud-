"""Post-process the pptxgenjs output:
1. remove mid-paragraph <a:pPr> (keeps bullets on lines with a bold lead-in);
2. add a formal Fade transition to every slide;
3. on the slide holding step_* shapes, reveal Scenario -> Think -> Decide for each case on click.
"""
import re, sys, zipfile, shutil

src = sys.argv[1]; tmp = src + ".tmp"
stray = re.compile(r'(</a:r>)<a:pPr\b[^>]*?(?:/>|>.*?</a:pPr>)', re.S)
TRANSITION = '<p:transition spd="med"><p:fade/></p:transition>'


def timing(steps):
    nid = [2]

    def n():
        nid[0] += 1
        return nid[0]

    clicks = []
    for spids in steps:
        effects = []
        for k, sp in enumerate(spids):
            node = "clickEffect" if k == 0 else "withEffect"
            effects.append(
                f'<p:par><p:cTn id="{n()}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{n()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f'<p:tgtEl><p:spTgt spid="{sp}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                f'<p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{n()}" dur="500"/>'
                f'<p:tgtEl><p:spTgt spid="{sp}"/></p:tgtEl></p:cBhvr></p:animEffect>'
                f'</p:childTnLst></p:cTn></p:par>')
        outer, inner = n(), n()
        clicks.append(
            f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            + "".join(effects) + '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    bld = "".join(f'<p:bldP spid="{sp}" grpId="0" animBg="1"/>' for st in steps for sp in st)
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            + "".join(clicks) +
            '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + bld + '</p:bldLst></p:timing>')


counts = {"stray": 0, "transitions": 0, "animated_slides": 0}
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if re.fullmatch(r"ppt/slides/slide\d+\.xml", item.filename):
            x = data.decode("utf-8")
            x, k = stray.subn(r"\1", x); counts["stray"] += k
            extra = TRANSITION
            ids = {name: int(i) for i, name in re.findall(r'<p:cNvPr id="(\d+)" name="(step_[^"]+)"', x)}
            if ids:
                steps = []
                for case in range(3):
                    for part in ("Scenario", "Think", "Decide"):
                        st = [ids[f"step_{case}_{part}_{e}"] for e in ("box", "lbl", "txt") if f"step_{case}_{part}_{e}" in ids]
                        if st:
                            steps.append(st)
                extra += timing(steps); counts["animated_slides"] += 1
            if "<p:transition" not in x:
                x = x.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + extra, 1); counts["transitions"] += 1
            data = x.encode("utf-8")
        zout.writestr(item, data)
shutil.move(tmp, src)
print(counts)
