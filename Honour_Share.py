"""
公共模塊程序庫. 重復使用的代碼在此更新....
"""
import numpy as np
import win32com.client
import os
import psutil
import configparser
import math
import pandas as pd
import clr
import xml.dom.minidom
import sys
from copy import copy
import datetime
import pymssql as sql



class master():
    def __init__(self,master_file,master_sheet):
        self.master = pd.read_excel(master_file, sheet_name=master_sheet)
        self.master_dict = copy(self.master)
        self.master_sheet_no_cn = "主檔無中文"
        self.gsm_na = "找不到克重!"
        for m in self.master:
            ms = self.master.get(m)
            # ms_notna = ms[ms.iloc[:, 1].notna()]  # 移去第1列為na 的數據. 是從0開始的.
            dict1 = {}
            for st in ms.itertuples():
                dict1[st[1]] = st[2]
                # print(st[0], st[1], st[2])
            self.master_dict[m] = dict1


    def get_dict(self,sheet,sheet_key,sheet_key_default="show find out info!"):
        """
        :param sheet:
        :param sheet_key:
        :param sheet_key_default: 默認找不到key 返回: sheet + ":" + str(sheet_key) + "在主檔找不到.", 否則就返回 sheet_key_default 內容
        :return:
        """
        if sheet_key:
            # print(self.master_dict.get(sheet, '').get(sheet_key, ""))
            # print("fanyi: ", sheet, sheet_key)# , sheet_value, type(sheet_value))
            if sheet_key_default=="show find out info!":
                sheet_key_default_value = sheet + ":" + str(sheet_key) + "在主檔找不到."
            else:
                sheet_key_default_value = sheet_key_default
            sheet_value=self.master_dict.get(sheet).get(sheet_key, sheet_key_default_value)

            if isinstance(sheet_value,float) or isinstance(sheet_value, int):
                if math.isnan(sheet_value):
                    sheet_value = sheet+":"+sheet_key+self.master_sheet_no_cn
            elif sheet_value==None:
                sheet_value = sheet+":"+sheet_key + self.master_sheet_no_cn
            # elif sheet=="Sample":
            #     sheet_value=sheet_value
            #     print(sheet_value)
            return sheet_value
        else:
            return ""

def Email_lotus(file,config):
    global key,df_all
    s = win32com.client.Dispatch('Notes.NotesSession')
    db_ddata=s.GetDatabase(config[1], r"PublicNSF\ddata.nsf")  #server , NSF path
    doc_ddata=db_ddata.GetDocumentByUNID("823BE41DAED99F2A48258810002FC9B5")
    st=config[5]+'To'
    ct = config[5] + 'CC'
    SendTo=doc_ddata.getitemvalue(st)
    CopyTo=doc_ddata.getitemvalue(ct)
    print("sendto ",SendTo)
    print("CopyTo ", CopyTo)
    db = s.GetDatabase(config[1], config[6])  #server , NSF path
    doc = db.CreateDocument
    doc.form = "Memo"

    # body1=doc.CreateRichTextItem("body")

    # tabs=np.array([["1","2"],["3","4"]])
    # tabs=[["1","2"],["3","4"]]
    style=s.CreateRichTextParagraphStyle
    style.LeftMargin =56
    style.FirstLineLeftMargin =56
    style.RightMargin = 1000

    # style[2].LeftMargin = 1
    # style[2].FirstLineLeftMargin = 1
    # style[2].RightMargin = 2
    body1 = doc.CreateRichTextItem("body")

    richStyle = s.CreateRichTextStyle
    # richStyle.NotesFont = 4
    richStyle.FontSize = 12
    richStyle.NotesFont = body1.GetNotesFont("Arial")
    body1.AppendStyle(richStyle)


    attachment = doc.CreateRichTextItem("Attachment")
    attachment.EmbedObject(1454, "", file, "補數檔案")
    doc.SendTo=SendTo
    doc.CopyTo = CopyTo


    datetime1=datetime.datetime.strftime(datetime.datetime.now(), "%Y-%m-%d %H:%M")
    subject=config[5] + ' '+datetime1+" 補數數據."
    doc.Subject=subject

    body=""


    for k in key:
        body=body+k+"\n"

    body1.Appendtext(body)
    head1="工单版本".ljust(8) + "貨號".ljust(10) + "識別碼".ljust(7) + "補數原因".ljust(18) + "編號".ljust(13)
    head2 = "JobVer".ljust(12) + "Item".ljust(12) + "Code".ljust(10) + "Reason".ljust(22) + "seq".ljust(20)
    body1.Appendtext(head1)
    body1.AddNewLine(1)
    body1.Appendtext(head2)
    body1.AddNewLine(1)

    for index, row in df_all.iterrows():
        if  pd.isna(row.iloc[5]):
            if type(row.iloc[4])==str:
                seq=row.iloc[4]
            else:
                seq=str(int(row.iloc[4]))
        else:
            seq=str(int(row.iloc[4]))+"-"+str(int(row.iloc[5]))
        rows=str(row.iloc[0]).ljust(12)+str(row.iloc[1]).ljust(12)+str(row.iloc[3]).ljust(10)+str(row.iloc[9]).ljust(20)+seq
        body1.Appendtext(rows)
        body1.AddNewLine(1)

    # doc.body=body

    # doc.save(True,False)
    doc.Send(False)

def get_Lotus_server(UserName):
    Lotus_server_dict = {
        "HMP": "HMP03/IT/HMP",
        "SHENGYI": "DG-ShengYi01/SHENGYI",
        "YAOHUI": "YaoHui01/IT/YAOHUI",
        "IndiaTeam": "IndiaTeam01/IndiaTeam",
        "MALAYSIATEAM": "Malaysia01/IT/MALAYSIATEAM",
        "HONOUR": "Honour01/HONOUR"
    }
    Server_key=str(UserName).split('=')[-1]
    Lotus_server=Lotus_server_dict.get(Server_key,"Honour01/HONOUR")
    print("server",Lotus_server)
    return Lotus_server

def Check_lotus_AppStore(app_name,FileVersion,Update):
    """
    檢查程序編號里的版本與要運行的程序版本是否一樣, 如果不一樣, 把Lotus 的下載出來.
    :param app_name:
    :param Lotus_server:
    :param FileVersion:
    :return:
    """
    # if Lotus_server=="":
    #     Lotus_server="HMP03/IT/HMP"
    s = win32com.client.Dispatch('Notes.NotesSession')
    Lotus_server = get_Lotus_server(s.UserName)
    db = s.GetDatabase(Lotus_server, r"PublicNSF\AppStore.nsf")
    if db.IsOpen and Update=="Yes":
        view = db.GetView("searchApp")
        doc = view.GetDocumentByKey(app_name, True)
        if doc is not None:
            Lotus_Version=doc.getitemvalue("Version")[0]
        if FileVersion!=Lotus_Version:
            rtitem = doc.GetFirstItem("programfiles")
            if rtitem is not None:
                for r in rtitem.EmbeddedObjects:
                    # print(r.Type, r.name)
                    rname = r.Name
                    file_type = os.path.splitext(rname)[-1].lower()
    
                    # HK_filename = os.getcwd()+"/Source/" + r.Name
                    filename=os.getcwd()+"/Source/" + r.Name
                    print("update file: ",filename)
                    r.ExtractFile(filename)
        return True
    elif db.IsOpen and Update!="Yes":
        return True
    else:
        return False

def get_lotus_AppStore():
    """
    獲取Lotus AppStore 模塊是Python 的程序編號.
    :param Lotus_server:
    :return: 程序編號
    """
    # if Lotus_server=="":
    #     Lotus_server="HMP03/IT/HMP"
    s = win32com.client.Dispatch('Notes.NotesSession')
    Lotus_server=get_Lotus_server(s.UserName)
    db = s.GetDatabase(Lotus_server, r"PublicNSF\AppStore.nsf")
    if db.IsOpen:
        view = db.GetView("searchAPP_AppUser")
        dc = view.GetAllDocumentsByKey(s.CommonUserName, True)
        # name=s.CommonUserName
    
        program_name = []
        program={}
        for i in range(1,dc.count+1):
            doc = dc.GetNthDocument(i)
            programid=doc.GetItemValue("programid")[0]
            programtitle=doc.GetItemValue("programtitle")[0]
            program[programid]=programtitle
            program_name.append(programid)
        return program,True
    else:
        return "",False

def kill_process(name):
    "結束程序運行."
    for proc in psutil.process_iter():
        # print(proc.name())
        if proc.name().startswith(name):
            proc.kill()


def update_ver(ver_txt):
    """
    自動更新版本號, 如果是EXE(frozen), 就不會更新版本號
    :param ver_txt: 版本txt文件
    :return:
    """
    # filevers 版本+1
    frozen = hasattr(sys, 'frozen')
    if not frozen:
        with open(ver_txt, 'r', encoding='utf-8') as file:
            file_contents = file.read()
            file.seek(0)
            for line in file:
                line=line.strip()
                if line.startswith('filevers='):
                    old_filevers = line
                    ver = line.split('=')[1]
                    ver1=ver[1:-2].split(',')
                    ver2 = [int(i) for i in ver1]
                    if ver2[3]>999:
                        ver2[3]=0
                        ver2[2] = ver2[2] + 1
                    if ver2[2]>999:
                        ver2[2] = 0
                        ver2[3] = 0
                        ver2[1] = ver2[1] + 1
                    if ver2[1]>999:
                        ver2[1] = 0
                        ver2[2] = 0
                        ver2[3] = 0
                        ver2[0] = ver2[0] + 1
                    else:
                        ver2[3] = ver2[3] + 1
                    new_filevers = "filevers=" + tuple(ver2).__str__()+","
                    print("文件版本:", new_filevers)
    
        #进行替换操作
        updated_contents = file_contents.replace(old_filevers, new_filevers)
    
        # 将修改后的内容写回文件（可选：可以写到一个新文件）
        with open(ver_txt, 'w', encoding='utf-8') as file:
            file.write(updated_contents)

def Py_Decrypto(userCommon_file, Honour_dll):
    """
    Honour.dll 加密,解密
    :param userCommon_file:.
    :param Honour_dll:
    :return: userID,userPWD,serverName,databaseName
    """

    # if not os.path.exists(r'C:\HonourProgram\Live\userCommon.xml'):
    #     shutil.copyfile('./Source/userCommon.xml',r'C:\HonourProgram\Live\userCommon.xml')
    #     print("update-UserCommon")
    # if not os.path.exists(r"C:\HonourProgram\Live\Appstore\Honour.dll"):
    #     shutil.copyfile('./Source/Honour.dll',r"C:\HonourProgram\Live\Appstore\Honour.dll")
    #     print("update-Honour.dll")
    # dom = xml.dom.minidom.parse(r'C:\HonourProgram\Live\userCommon.xml')
    dom = xml.dom.minidom.parse(userCommon_file)
    root = dom.documentElement  # students结点
    # print(root.getAttribute('userCommon'))
    userIDSQL = root.getElementsByTagName('userIDSQL')
    userPWDSQL=root.getElementsByTagName('userPWDSQL')

    serverName = root.getElementsByTagName('serverName')[0].firstChild.data
    databaseName = root.getElementsByTagName('databaseName')[0].firstChild.data

    userID_Encrypt=userIDSQL[0].firstChild.data
    userPWD_Encrypt=userPWDSQL[0].firstChild.data

    # clr.AddReference(r"C:\HonourProgram\Live\Appstore\Honour.dll")
    clr.AddReference(Honour_dll)

    # import the namespace and class

    from Honour import SymmetricMethod

    # create an object of the class

    obj = SymmetricMethod()

    v1=obj.Encrypto("beginer")
    # print("加密: ",v1)
    value = obj.Decrypto("ivTrAqqgiIMUD4RL31nRgA==")
    # print("解密: ", value)

    userID=obj.Decrypto(userID_Encrypt)
    userPWD = obj.Decrypto(userPWD_Encrypt)
    # print("解密: ", userID,userPWD)

    return userID,userPWD,serverName,databaseName

def read_sql_fetchall(sql_str,userID, userPWD, serverName, databaseName):


    conn = sql.connect(server=serverName, user=userID, password=userPWD, database=databaseName,tds_version="7.0")
    # stock_basic = conn.cursor(as_dict=True)

    stock_basic = conn.cursor()
    # stock_basic.execute("select  max(id) as num from [dbo].[mv_job] where jobver=%s",jobver)

    # sql_str=f"SELECT templateCode,suffix,widthmm1,widthmm2,UnitArea FROM [HMPSQL01].[dbo].[V_StdTemplate]  where suffix='{Suffix}' and ((widthmm1={m_long} and widthmm2={m_width}) or (widthmm1={m_width} and widthmm2={m_long})) order by suffix,widthmm1,widthmm2,UnitArea desc"
    stock_basic.execute(sql_str)

    #data=stock_basic.fetchone()
    all=stock_basic.fetchall()
    # print(all)
    # max_id = stock_basic.fetchone()
    # if max_id is None:
    #     max_id=0
    # jobver_id = jobver + "-" + str(max_id + 1).rjust(4, '0')
    # print(max_id)
    # values1=(max_id + 1,jobver,jobver_id,process_department)
    # stock_basic.execute("insert into  [dbo].[mv_job] ([id],[jobver],[jobver_id],[dept]) values(%d,%s,%s,%s)", values1)
    conn.commit()
    return all
    # print(jobver_id)

if __name__ == '__main__':
    # userID, userPWD, serverName, databaseName = Py_Decrypto(r'C:\HonourProgram\Live\userCommon.xml', r"C:\HonourProgram\Live\Appstore\Honour.dll")
    # print(userID,userPWD)
    # # update_ver("./ver1.txt")
    # master_file = r"D:\xu\Python\Check_Innsight/master/master.xlsx"
    # master_sheet = ["stockType", "Folding", "imagingOption", "personalizationOption", "numberOfPages", "numSheet", "perforation", "sealaffixedMaterial", "Sample", "ProductionReport_Head"]
    # master1 = master(master_file, master_sheet)
    # t:dict=get_lotus_AppStore()
    # userID, userPWD, serverName, databaseName = Py_Decrypto(r'C:\HonourProgram\Live\userCommon.xml', r"C:\HonourProgram\Live\Appstore\Honour.dll")
    key=["abc","123"]
    file=r"D:\xu\Python\QtPDF\bs_file1\bs_template_ok.xlsx"
    df_all=pd.read_excel(file)
    config=('ITProg2', 'HMP03/IT/HMP', 'PublicNSF\\QM.nsf', None, 'bs_file', 'ITProg2', 'mail\\ITProg2.nsf')
    Email_lotus(file,config)
    print("test1:")