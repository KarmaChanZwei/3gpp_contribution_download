# 3gpp_contribution_download

For CMCC 3GPP RAN2 Team

Download contributions on 3GPP server listed in "TDoc_List_XXXX.xlsx", archive and rename them, useful in pre-meeting Tdoc review.

dw2.py will rename the Tdoc to "TDocNo Source". dw3.py will rename the Tdoc to "TDocNo Source Title", with filename check for windows. dw3_trun.py is the same with dw3.py but will truncate filename longer than MAX_PATH for windows.

If you don't know MAX_PATH, it is recommended to use dw2.py and dw3_trun.py. Anyway, it's your choice.

How to use:
  1. Install Windows Terminal, https://learn.microsoft.com/zh-cn/windows/terminal/.
  2. Install Python, recommended 3.11.9, [https://www.python.org/downloads/](https://www.python.org/downloads/release/python-3119/). Use "python --version" to check version. Remember to add python.exe to PATH.
  3. Run "pip install pandas openpyxl requests" in Terminal. If you have network issue, use "pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pandas openpyxl requests" instead. 
  4. Copy "TDoc_List_XXXX.xlsx" and scripts to an empty folder, right click on blank space and choose "Open in the terminal", then use "python dw2.py" , "python dw3.py", or "python dw3_trun.py". Script will automaticly creat agenda item folder and TDocs will be automaticly archived.
  5. (Not necessary) run "unzip.py" in terminal to extract all download contributions. 

If there are multiple "TDoc_List_XXXX.xlsx", scripts will only process the first one.
If you only need certain agenda items, please edit the .xlsx and KEEP TABLE HEADER.

<img width="912" alt="2a978c2c3886fbdf5301037012dd5d8" src="https://github.com/user-attachments/assets/b9718a3b-fbbc-4631-b55d-c93c4635eb81">
