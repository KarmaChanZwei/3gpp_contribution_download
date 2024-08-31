# 3gpp_contribution_download

Download contributions on 3GPP server listed in TDic_List_XXXX.xlsx, archive and rename them, useful in pre-meeting Tdoc review.

How to use:
  1. Install Windows Terminal, https://learn.microsoft.com/zh-cn/windows/terminal/.
  2. Install Python, recommended 3.11.9, [https://www.python.org/downloads/](https://www.python.org/downloads/release/python-3119/). Use "python --version" to check version. Remember to add python.exe to PATH.
  3. Run "pip install pandas openpyxl requests" in Terminal.
  4. Copy TDoc_List_XXXX.xlsx and python script to an empty folder, run it in Terminal, use "python dw2.py" or "python dw3.py".

Suggest to creat an empty folder, and put the TDoc_List_XXXX.xlsx together with the python scripts. Script will automaticly creat new folder with agenda item. And TDocs of that agenda will be automaticly archived in that folder. 

If there are multiple .xlsx, scripts will only process the first one.
If you only need certain agenda items, please edit the .xlsx and KEEP TABLE HEADER.

dw2.py will rename the Tdoc to "TDocNo Source". dw3.py will rename the Tdoc to "TDocNo Source Title", with filename check for windows.

NOTE: Windows has a 260 character limitation on MAX_PATH. It is recommended to use dw2.py and dw3_trun.py if TDoc is co-signed with multiple companies. Anyway, it's your choice.

<img width="912" alt="2a978c2c3886fbdf5301037012dd5d8" src="https://github.com/user-attachments/assets/b9718a3b-fbbc-4631-b55d-c93c4635eb81">


For CMCC RAN2 Team
