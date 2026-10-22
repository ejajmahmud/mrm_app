from pathlib import Path
import zipfile
ROOT=Path(__file__).resolve().parents[1]
files=['MRM_Complete_Connected_Platform_96_Pages.pptx','MRM_Complete_Connected_Platform_96_Pages.pdf','MRM_User_Value_Bridges_24_Pages.pptx','MRM_User_Value_Bridges_24_Pages.pdf','README.md','Verification.json','Bridge_Model.json']
with zipfile.ZipFile(ROOT/'MRM_User_Journeys_Download.zip','w',zipfile.ZIP_DEFLATED) as z:
 for name in files:z.write(ROOT/name,name)
with zipfile.ZipFile(ROOT/'MRM_User_Journeys_Download.zip') as z:assert z.testzip() is None
print('Bundle MiB:',round((ROOT/'MRM_User_Journeys_Download.zip').stat().st_size/1024/1024,1))
