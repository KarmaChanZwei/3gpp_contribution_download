# 3gpp_contribution_download

Download contributions on 3GPP server listed in TDic_List_XXXX.xlsx, archive and rename them, useful in pre-meeting Tdoc review.

Before use：
  0. Install Windows Terminal, https://learn.microsoft.com/zh-cn/windows/terminal/ .
  1. Download Python, recommended 3.11.9, [https://www.python.org/downloads/](https://www.python.org/downloads/release/python-3119/). Use "python --version" to check. Rememeber to add python.exe to PATH.
  2. Run "pip install pandas openpyxl requests"
  3. Copy TDoc_List_XXXX.xlsx and python script to an empty folder, run it in terminal, "python dw3.py" or "python dw3.py".

Suggest to creat an empty folder, and put the TDoc_List_XXXX.xlsx toegther with the python scripts. Script will automaticly creat new folder with agenda item. And TDocs of that agenda will be automaticly archived in that folder. 

If there are multiple .xlsx, scripts will only use the first one.
If you only need certain ageneda items, please edit the .xlsx and keep table header.

dw2.py will rename the Tdoc to "TDocNo Source". dw3.py will rename the Tdoc to "TDocNo Source Title", with filename check for windows.

<img width="912" alt="2a978c2c3886fbdf5301037012dd5d8" src="https://github.com/user-attachments/assets/b9718a3b-fbbc-4631-b55d-c93c4635eb81">


CMCC RAN2 Team
