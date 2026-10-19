"""Create the public-safe screenshot register from corrected source metadata.
OCR text and the identity number in withheld files are never published.
"""
from pathlib import Path
import json,ast,hashlib,os
from PIL import Image
ROOT=Path(__file__).resolve().parents[1]
SCRATCH=Path('/workspace/scratch/mrm-rebuild')
body=ast.parse((ROOT/'source'/'build_deck.py').read_text())
CAPTIONS=next(ast.literal_eval(n.value) for n in body.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='CAPTIONS' for t in n.targets))
src=Path(os.environ.get('MRM_SCREENSHOT_DIR',str(SCRATCH/'corrected-screens')))
rows=[]
for p in sorted(src.glob('*.png')):
 im=Image.open(p).convert('RGB').crop((150,180,1320,2800))
 rows.append({'file':p.name,'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'content_sha256':hashlib.sha256(im.tobytes()).hexdigest()})
assert len(rows)==48,'Supply all 48 corrected screenshots.'
uses=json.loads((ROOT/'Image_Use_Register.json').read_text());pages={x['screen']:x['page'] for x in uses}
findings={
'01':'Brand/loading state; verification, source-sync and audit language is not independently validated.',
'02':'Onboarding pillar 1: care and mentorship. Service hours, shelter and kit promises need operating approval.',
'03':'Welcome / guest entry. Volunteer counts, aid totals and audit badges are illustrative.',
'04':'Track selection at registration is not completed volunteer / teacher approval.',
'05':'Sign-in, recovery and biometric controls are visible; authentication and biometric execution were not tested.',
'06':'Admin Ops role state with utilities; not the ordinary member home or proof of live operations.',
'07':'Utility menu also includes development helpers. The framework text does not establish delivered architecture.',
'08':'Analytics plus income/savings estimate; not case-tracking evidence. Amounts, targets and authority wording unverified.',
'09':'Search explicitly includes Quran, duas, ingredient references and mosques. Not proof of approved volunteer-map access.',
'10':'Relief campaign list, not retail marketplace. Fundraising, tax and impact claims require approval.',
'11':'Campaign Overview tab, not a commercial product checkout. Delivery and authority claims unverified.',
'12':'Profile and six demo personas; server role approval and clearance are not demonstrated.',
'13':'Settings include zone, adhan, cache, disguise and lock controls. Toggles do not prove their behaviour.',
'14':'Calendar combines programme and duty slots; it does not prove registered attendance or operational delivery.',
'15':'Community feed with example dispatches and stories. Outcomes and implied verification are not evidence.',
'16':'Assigned caseworker conversation; encryption, voice playback, calls and ETA claims are not tested.',
'17':'Sadaqah/Infaq purpose allocation, not a retail cart. Example total is not a teacher plan price.',
'18':'Withheld: payment form contains a full identity number plus personal details. Gateway and tax claims unverified.',
'19':'In-app notifications. Tax receipt and dispatch examples are not live-event evidence or safe lock-screen payloads.',
'20':'Loading is the selected state; this file alone does not demonstrate empty/error behaviour.',
'21':'Waqt/fasting/prayer-window view; time conventions, religious guidance and source-sync wording need review.',
'22':'Private personal Salah log with sample counts. It must not become a public faith score or volunteer rank.',
'23':'A simulation slider is visible. A compass drawing and distance list do not validate GPS, sensors or 3D navigation.',
'24':'Reader shows Al-Fatihah while a Ya-Sin tab says Selected. Describe the visible verses and resolve tab consistency.',
'25':'Selected phrase, counter, lap and total display are visible; counter persistence and semantics need testing.',
'26':'Dua categories with references and recitation controls; a cited hadith does not demonstrate a separate Hadith reader.',
'27':'Only posture 1 of 7 is shown. Other postures and pronunciation behaviour were not inspected in runtime.',
'28':'Asma-ul-Husna names and meanings, not a Hadith reader or proof of source approval.',
'29':'Two functions share the view: Fitrah/Fidyah estimate and ingredient-reference list. Rates and classifications need review.',
'30':'Volunteer on-duty role view; GPS Verified and dispatch indicators are sample UI, not tested matching or location.',
'31':'Seeker role view; care-line availability, complimentary kits and confidentiality require operating and technical checks.',
'32':'Scholar/Mentor roster with Start Room control; not a subscription, teacher-approval or working video-conference screen.',
'33':'Applicant progress includes police-clearance wording; agreed document review and interview remain separate policy requirements.',
'34':'Member role view retains utilities and a sample campaign. Broad Muslim access is retained without a teacher paywall.',
'35':'Onboarding pillar 2 describes welfare, fundraising and tax claims; it is not a payment form or authority approval.',
'36':'Onboarding pillar 3 describes volunteering. Dispatch, source accuracy and encryption claims remain unverified.',
'37':'Provisions tab is distinct from Overview; it lists intended kit contents, not verified delivery.',
'38':'Impact tab has sample named testimonials and Verified badges. No testimonial was independently confirmed.',
'39':'Calendar almanac is a distinct selected tab with prayer/fasting detail, not a duplicate of the booking list.',
'40':'Slot-confirmation modal is a UI state, not proof of persistent registration or attendance.',
'41':'Community composer is visible; no independent moderation queue, publication permissions or storage is demonstrated.',
'42':'Withheld: receipt concept includes the same full identity number and personal details. No transfer or lawful receipt verified.',
'43':'Empty panel is visible. Wording that all dispatches are complete cannot be inferred from an empty query.',
'44':'Network-error state is visible. Cached Quran/prayer availability remains a claim requiring a real offline test.',
'45':'Role-picker modal is distinct from Profile; changing a persona must not grant server permissions.',
'46':'Directory lists 20 base screens. The 48 files include role, tab, modal and theme states, not 48 independent modules.',
'47':'Excluded from feature imagery: development installation instructions. Not evidence of native release or framework approval.',
'48':'A distinct dark-theme Home state. Used to explain theme preference, not counted as another operational module.'}
register=[]
for r in rows:
 sid=r['file'][:2]
 register.append({'id':sid,'file':r['file'],'actual_visible_content':CAPTIONS[sid],'main_deck_page':pages.get(sid),'disposition':'Used once' if sid in pages else 'Excluded','finding':findings[sid],'source_sha256':r['sha256'],'content_region_sha256':r['content_sha256']})
assert len(register)==48 and len(pages)==45
assert len(set(r['source_sha256'] for r in register))==48
assert len(set(r['content_region_sha256'] for r in register))==48
(ROOT/'Screen_Audit.json').write_text(json.dumps(register,indent=2))
lines=['# Corrected screenshot audit','',
'All 48 images from the corrected Drive folder were visually inspected and OCR-reviewed. The withdrawn ZIP was not used. Exact byte hashes and a consistent interior-content crop found no exact duplicates. Tab, role, modal and theme variations were checked semantically; they are not counted as separate independent modules.','',
'The main deck uses 45 distinct captures once each. Payment form 18 and receipt 42 are withheld because they contain a full identity number and personal details. Capture 47 is development installation help. Source PNGs are retained outside the repository; display copies are resized to 900 pixels wide with original proportions and transparency for a practical high-resolution deck size. Screenshot content is not rewritten.','',
'Operational statistics, tax/audit badges, testimonials, certification labels, encryption claims and source-sync wording in screenshots remain illustrative or unverified. Captures are UI-state evidence; they do not establish working APIs, persistence, payment, authority approval or production performance.','',
'| ID | Actual visible content | Deck page | Review finding |','| --- | --- | --- | --- |']
for r in register:lines.append(f"| {r['id']} | {r['actual_visible_content']} | {r['main_deck_page'] or 'Excluded'} | {r['finding']} |")
(ROOT/'Screen_Audit.md').write_text('\n'.join(lines)+'\n')
print('Audited 48 corrected images; 45 unique uses; 3 justified exclusions; no exact image/content duplicates.')
