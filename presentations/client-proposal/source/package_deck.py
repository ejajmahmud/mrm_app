from pathlib import Path
import zipfile
ROOT=Path(__file__).resolve().parents[1]
files=['MRM_Complete_Client_Proposal.pptx','MRM_Complete_Client_Proposal.pdf','README.md','Verification.json','previews/Cover.png']
with zipfile.ZipFile(ROOT/'MRM_Client_Proposal_Download.zip','w',zipfile.ZIP_DEFLATED) as z:
 for name in files:z.write(ROOT/name,name)
with zipfile.ZipFile(ROOT/'MRM_Client_Proposal_Download.zip') as z:assert z.testzip() is None
print('Bundle MiB:',round((ROOT/'MRM_Client_Proposal_Download.zip').stat().st_size/1024/1024,1))
