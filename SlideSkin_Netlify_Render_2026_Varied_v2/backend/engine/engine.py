"""Ray's Magic PowerPoint Converter: geometry-aware deterministic beta.
Uses native 2026 slide designs. Preserves editable text, raster images and notes.
Reports unsupported native objects instead of claiming perfect preservation.
"""
import io,json,re,math
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from base_engine import clone_template_slide,set_text,EMU

INTRO=8; OUTRO=18
TEXT_CANDIDATES=(11,12,15,16)
IMAGE_CANDIDATES=(9,11,12)

def extract(slide):
    shapes=[s for s in slide.shapes if s.has_text_frame and s.text.strip() and not re.fullmatch(r'\d+',s.text.strip())]
    title_shape=slide.shapes.title
    title=title_shape.text.strip() if title_shape is not None and title_shape.text.strip() else (shapes[0].text.strip() if shapes else 'Untitled')
    rest=[s.text.strip() for s in shapes if s is not title_shape and s.text.strip()!=title]
    # A paragraph of prose is not a subtitle. Keep it in the body.
    subtitle=rest.pop(0) if rest and len(rest[0])<=110 and '\n' not in rest[0] and '\v' not in rest[0] else ''
    lines=[line.strip().lstrip('•').strip() for block in rest for line in re.split(r'[\n\v]',block) if line.strip()]
    return dict(title=title,subtitle=subtitle,lines=lines,images=[s for s in slide.shapes if s.shape_type==MSO_SHAPE_TYPE.PICTURE],
       charts=sum(bool(s.has_chart) for s in slide.shapes),tables=sum(bool(s.has_table) for s in slide.shapes),
       media=sum(s.shape_type==MSO_SHAPE_TYPE.MEDIA for s in slide.shapes))

def kind(c):
    t=c['title'].lower().strip()
    if re.match(r'^(thank you|thanks|questions|q\s*&\s*a|the end|closing)\b',t):return 'closing'
    if re.match(r'^(welcome|introduction|overview|opening|title slide)\b',t):return 'intro'
    return 'content'

def layout(c,usage,prev):
    n=len(c['lines']); size=sum(map(len,c['lines']))+len(c['subtitle']); imgs=len(c['images'])
    options=IMAGE_CANDIDATES if imgs else TEXT_CANDIDATES
    scores=[]
    for i in options:
        # 9: left text/right visual; 11: full width; 12: two columns; 15/16: three columns
        if imgs:
            base={9:9,11:5,12:8}[i]
            if size>440:base+= {9:-3,11:2,12:-1}[i]
            if imgs>1:base+= {9:-2,11:1,12:2}[i]
        else:
            base={11:7,12:6,15:5,16:5}[i]
            if size>600:base+= {11:4,12:2,15:-5,16:-5}[i]
            if 3<=n<=7 and size<550:base+= {11:0,12:2,15:4,16:4}[i]
            if n<=2:base+= {11:4,12:-2,15:-4,16:-4}[i]
        base-=min(usage.get(i,0),5)*0.7
        if prev==i:base-=5
        scores.append((base,i))
    return max(scores)[1]

def text_slots(slide):
    bodies=[]; heads=[]
    for s in slide.shapes:
        if not s.has_text_frame:continue
        name=s.name.lower()
        if 'content placeholder' in name:
            (heads if s.height<.75*EMU else bodies).append(s)
    return sorted(heads,key=lambda s:s.left),sorted(bodies,key=lambda s:s.left)

def remove_demo(slide):
    # Clear ALL example text, including unusual text shapes in slide 17, without touching decorative graphics.
    for s in slide.shapes:
        if not s.has_text_frame:continue
        n=s.name.lower()
        if 'title' in n and 'subtitle' not in n:continue
        if 'slide number' in n:continue
        if s.text.strip():set_text(s,'')

def fit_image(slide,im,region):
    from PIL import Image
    blob=im.image.blob
    with Image.open(io.BytesIO(blob)) as img:iw,ih=img.size
    x,y,w,h=region
    ratio=min(w/iw,h/ih)
    ww=max(1,int(iw*ratio));hh=max(1,int(ih*ratio))
    slide.shapes.add_picture(io.BytesIO(blob),x+(w-ww)//2,y+(h-hh)//2,width=ww,height=hh)

def fill_images(slide,images,region):
    if not images:return
    x,y,w,h=region
    count=len(images);cols=1 if count==1 else 2;rows=math.ceil(count/cols);gap=int(.10*EMU)
    cw=(w-gap*(cols-1))//cols;ch=(h-gap*(rows-1))//rows
    for j,im in enumerate(images):fit_image(slide,im,(x+(j%cols)*(cw+gap),y+(j//cols)*(ch+gap),cw,ch))

def distribute(lines,slots):
    if not slots:return []
    # Fill by estimated usable character capacity rather than equal line counts.
    caps=[max(1,int((s.width/EMU)*(s.height/EMU)*12)) for s in slots]
    groups=[[] for _ in slots];idx=0;used=0
    for line in lines:
        if idx<len(slots)-1 and used+len(line)>caps[idx] and groups[idx]:idx+=1;used=0
        groups[idx].append(line);used+=len(line)
    return groups

def apply(slide,c,page,index,source):
    warnings=[]
    title=next((s for s in slide.shapes if s.has_text_frame and 'title' in s.name.lower() and 'subtitle' not in s.name.lower()),None)
    remove_demo(slide)
    if title:set_text(title,c['title'])
    for s in slide.shapes:
        if s.has_text_frame and 'slide number' in s.name.lower():set_text(s,str(page))
    heads,bodies=text_slots(slide)
    if c['images']:
        # Native template content regions are the alignment anchors.
        if index==9:
            region=(int(6.6*EMU),int(1.7*EMU),int(5.9*EMU),int(4.75*EMU))
        elif index==12:
            region=(int(6.7*EMU),int(1.9*EMU),int(5.7*EMU),int(4.3*EMU))
            bodies=[s for s in bodies if s.left<6.5*EMU]
            heads=[s for s in heads if s.left<6.5*EMU]
        else:
            region=(int(7.2*EMU),int(1.9*EMU),int(5.1*EMU),int(4.3*EMU))
            for s in bodies:
                if s.left+s.width>region[0]-int(.2*EMU):s.width=max(int(2.4*EMU),region[0]-int(.35*EMU)-s.left)
            for s in heads:
                if s.left+s.width>region[0]-int(.2*EMU):s.width=max(int(2.4*EMU),region[0]-int(.35*EMU)-s.left)
        fill_images(slide,c['images'],region)
    # Headings are aligned to template's native horizontal rule/heading strip at y=1.5.
    if heads and c['subtitle']:set_text(heads[0],c['subtitle'],16)
    elif c['subtitle']:
        # No subhead placeholder: include subtitle in body, never float a new textbox over the rule.
        c=dict(c,lines=[c['subtitle']]+c['lines'])
    if index in (15,16) and not c['images'] and len(bodies)>1:
        groups=distribute(c['lines'],bodies)
        for j,(s,g) in enumerate(zip(bodies,groups)):
            set_text(s,'\n'.join('• '+v for v in g),16 if sum(map(len,g))>180 else 18)
            if j<len(heads):set_text(heads[j],c['subtitle'] if j==0 and c['subtitle'] else f'POINT {j+1}',15)
    elif bodies:
        target=max(bodies,key=lambda s:s.width*s.height)
        lines=c['lines']
        if index==12 and not c['images'] and len(bodies)>=2:
            groups=distribute(lines,bodies)
            for s,g in zip(bodies,groups):set_text(s,'\n'.join('• '+v for v in g),17)
        else:
            size=18 if sum(map(len,lines))<340 else 16
            set_text(target,'\n'.join('• '+v for v in lines),size)
            cap=max(1,int(target.width/EMU*target.height/EMU*12))
            if sum(map(len,lines))>cap:warnings.append('Text may overflow; review this slide')
    elif c['lines']:
        warnings.append('No native text region available; content requires review')
    for key in ('charts','tables','media'):
        if c[key]:warnings.append(f'{c[key]} source {key} object(s) require manual recreation; not preserved')
    try:
        note=source.notes_slide.notes_text_frame.text
        if note:slide.notes_slide.notes_text_frame.text=note
    except Exception:pass
    return warnings

def cover(deck,originals,index,title,subtitle,source=None):
    dst=clone_template_slide(deck,originals[index-1])
    for s in dst.shapes:
        if not s.has_text_frame:continue
        name=s.name.lower()
        if 'subtitle' in name:set_text(s,subtitle)
        elif 'title' in name:set_text(s,title)
    if source is not None:
        try:dst.notes_slide.notes_text_frame.text=source.notes_slide.notes_text_frame.text
        except Exception:pass
    return dst

def run(template,source,output,report_path=None,min_pt=16):
    deck=Presentation(str(template));incoming=Presentation(str(source));originals=list(deck.slides)
    if len(originals)<19:raise ValueError(f'The 2026 template requires the 19-slide reference deck; found {len(originals)}')
    contents=[extract(s) for s in incoming.slides]
    report=[];usage={};prev=None;page=0
    # Always create explicit bookends. Existing source opening/closing content is retained too.
    intro=contents[0]['title'] if contents else 'Presentation'
    cover(deck,originals,INTRO,intro,'Ray\'s Magic PowerPoint Converter')
    page+=1;report.append(dict(source_slide=None,output_slides=[page],template_slide=INTRO,mode='generated_introduction',warnings=[]))
    for i,(src,c) in enumerate(zip(incoming.slides,contents),1):
        idx=layout(c,usage,prev)
        usage[idx]=usage.get(idx,0)+1;prev=idx
        dst=clone_template_slide(deck,originals[idx-1]);page+=1
        warnings=apply(dst,c,page,idx,src)
        report.append(dict(source_slide=i,output_slides=[page],template_slide=idx,mode='content_matched',warnings=warnings))
    cover(deck,originals,OUTRO,'Thank You!','Questions?')
    page+=1;report.append(dict(source_slide=None,output_slides=[page],template_slide=OUTRO,mode='generated_closing',warnings=[]))
    for _ in originals:
        sid=deck.slides._sldIdLst[0];deck.part.drop_rel(sid.rId);deck.slides._sldIdLst.remove(sid)
    deck.save(str(output))
    if report_path:Path(report_path).write_text(json.dumps(report,indent=2))
    return report
