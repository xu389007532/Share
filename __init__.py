"""
功能: 自動把Base環環里的模塊更新到其它虛擬環境里. 并清空其它虛擬環境的__init__.py 文件.
__init__.py文件在其它虛擬環境不需要寫內容. 只在Base環境.
source_path: 是主路徑, 全部數據在此更新
update_path: 當有新的虛擬環境, 需要加入.
update_file: 如果有多個py文件, 需要加入.
"""
import os.path
from shutil import copyfile
name=os.environ.get("USERNAME")
if os.environ.get("USERNAME")=='ITProg02':
    source_path=r"C:\Users\ITProg02\AppData\Local\anaconda3\envs\py3.12\Lib\site-packages\Share"
    update_path=[r"C:\Users\ITProg02\AppData\Local\anaconda3\envs\inputbs\Lib\site-packages\Share", r"C:\Users\ITProg02\AppData\Local\anaconda3\envs\DIFS\Lib\site-packages\Share"]
    update_file=["Honour_Share.py"]
    clear_init_file = "__init__.py"
    for up in update_path:
        for uf in  update_file:
            source_file_path = source_path + "/" + uf
            update_file_path=up+"/"+uf
            source_time=os.path.getmtime(source_file_path)
            update_time = os.path.getmtime(update_file_path)
            if source_time>update_time:
                copyfile(source_file_path,update_file_path)
                print("Share 包內模塊已更新: ",source_file_path,update_file_path)
        init_file=up+"/"+clear_init_file
        with open(init_file, 'r+', encoding='utf-8') as file:
            check=file.read()
            if check!="#不是Base 環境, Share 包的__init__.py文件不要有內容.":
                file.write("#不是Base 環境, Share 包的__init__.py文件不要有內容.")
                print("包:"+up+"/__init__.py 已更新")
            else:
                print("包:"+up+"/__init__.py 不需要更新")

