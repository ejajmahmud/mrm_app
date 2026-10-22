from pathlib import Path
import json,re,zipfile,hashlib
import fitz
from PIL import Image,ImageDraw
from pptx import Presentation
ROOT=Path(__file__).resolve().parents[1];STEM='MRM_Complete_Client_Proposal'
N=lambda s:re.sub('[^a-z0-9]','',s.lower())
ppt=Presentation(ROOT/(STEM+'.pptx'));doc=fitz.open(ROOT/(STEM+'.pdf'));assert len(ppt.slides)==len(doc)==85
missing=[];bounds=[];links=0;ids={s.slide_id:i for i,s in enumerate(ppt.slides)}
for pg in doc:
 for a in pg.get_links():pg.delete_link(a)
for i,(sl,pg) in enumerate(zip(ppt.slides,doc)):
 content=N(pg.get_text())
 for sh in sl.shapes:
  if sh.has_text_frame and len(N(sh.text))>1 and N(sh.text) not in content:missing.append([i+1,sh.text])
  if sh.left<0 or sh.top<0 or sh.left+sh.width>ppt.slide_width+10 or sh.top+sh.height>ppt.slide_height+10:bounds.append([i+1,sh.name])
  # All new opening and closing text should stay inside its own textbox.
  if (i<5 or i>=83) and sh.has_text_frame and N(sh.text):
   r=fitz.Rect(sh.left/12700-1,sh.top/12700-1,(sh.left+sh.width)/12700+1,(sh.top+sh.height)/12700+1)
   if N(sh.text) not in N(pg.get_textbox(r)):bounds.append([i+1,'textbox',sh.text])
  target=sh.click_action.target_slide
  if target is not None:
   assert target.slide_id in ids
   rect=fitz.Rect(sh.left/12700,sh.top/12700,(sh.left+sh.width)/12700,(sh.top+sh.height)/12700)
   pg.insert_link({'kind':fitz.LINK_GOTO,'from':rect,'page':ids[target.slide_id],'to':fitz.Point(0,0)});links+=1
 for block in pg.get_text('dict')['blocks']:
  if block['type']!=0:continue
  for line in block['lines']:
   for span in line['spans']:
    x0,y0,x1,y1=span['bbox']
    if min(x0,y0)<-.5 or x1>pg.rect.width+.5 or y1>pg.rect.height+.5:bounds.append([i+1,'rendered span',span['text']])
assert not missing,missing
assert not bounds,bounds
# Cover actually embeds three separate app screenshots.
cover=ppt.slides[0];photos=[s for s in cover.shapes if s.shape_type==13 and s.height>ppt.slide_height/2]
assert len(photos)==3 and len({hashlib.sha256(s.image.blob).hexdigest() for s in photos})==3
# Agenda destinations, source remapping and page labels must match the new order.
for i,s in enumerate(ppt.slides):
 if i:assert any(sh.has_text_frame and sh.text==f'{i+1:02d} / 85' for sh in s.shapes)
assert not any(' / 72' in s.text or ' / 24' in s.text for slide in ppt.slides for s in slide.shapes if s.has_text_frame)
titles=json.loads((ROOT/'Contents.json').read_text());doc.set_toc([[1,t,i+1] for i,t in enumerate(titles)])
temp=ROOT/(STEM+'_linked.pdf');doc.save(temp,garbage=4,deflate=True);doc.close();temp.replace(ROOT/(STEM+'.pdf'))
doc=fitz.open(ROOT/(STEM+'.pdf'));assert sum(len(p.get_links()) for p in doc)==links
pdir=ROOT/'previews';pdir.mkdir(exist_ok=True)
for start in range(0,85,6):
 sheet=Image.new('RGB',(1800,1650),'#111B35');d=ImageDraw.Draw(sheet)
 for j in range(start,min(start+6,85)):
  pix=doc[j].get_pixmap(matrix=fitz.Matrix(.65,.65));im=Image.frombytes('RGB',[pix.width,pix.height],pix.samples);im.thumbnail((890,501));x=(j-start)%2*900;y=(j-start)//2*550;sheet.paste(im,(x,y+25));d.text((x+10,y+5),f'PAGE {j+1:02d}',fill='white')
 sheet.save(pdir/f'Pages_{start+1:02d}_{min(start+6,85):02d}.jpg')
for i,name in [(0,'Cover'),(1,'Executive_Summary'),(3,'Connected_Platform'),(4,'Proposal_Guide'),(7,'Teacher_Volunteer_Bridge'),(84,'Closing')]:
 doc[i].get_pixmap(matrix=fitz.Matrix(1.3334,1.3334)).save(pdir/(name+'.png'))
# Record intended screenshot repetition transparently: cover hero images recur in feature detail.
assembly=json.loads((ROOT/'Assembly_Model.json').read_text());mapping={row['source_page']:row['page'] for row in assembly if row['source']=='original'}
uses=json.loads((ROOT.parent/'strategic-rebuild/Image_Use_Register.json').read_text());register=[]
for row in uses:
 if row['page'] in mapping:register.append(dict(row,page=mapping[row['page']],purpose='feature explanation'))
for sid in ['24','30','34']:register.append({'page':1,'screen':sid,'purpose':'intentional cover hero; also explained in the feature section'})
(ROOT/'Image_Use_Register.json').write_text(json.dumps(register,indent=2))
checks={'pages':85,'internal_navigation_links':links,'missing_native_text':0,'out_of_bounds':0,'opening_closing_textbox_overflows':0,'cover_distinct_app_screens':3,'global_page_numbers':'passed','editable_slide_links':'passed','screenshots':{'unique_sources':len({r['screen'] for r in register}),'placements':len(register),'intentional_cover_reuses':3},'structure':'Cover → proposition → audience → connected experience → guide → user journeys → full feature design → service systems → extension concepts → delivery → scope → next steps'}
(ROOT/'Verification.json').write_text(json.dumps(checks,indent=2));print(json.dumps(checks,indent=2))
