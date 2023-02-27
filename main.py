import os
import time

from pywinauto.application import Application
from pywinauto.timings import Timings
import win32con
import win32gui

# 运行前需要将Pdg2Pic参数设置第二个选项“目标文件夹与源文件夹相同”勾选上

N = 4  # 进程数量
Timings.fast()


# 迭代获取out_folder内部所有含有pdg的文件夹
def get_pdg_folders(out_folder):
    temp_folders = [out_folder]
    pdg_folders = []
    while len(temp_folders) > 0:
        temp_folder = temp_folders.pop()
        # 是否直接包含pdg
        has_pdg = False
        for f in os.scandir(temp_folder):
            if f.is_file() and f.path.endswith(".pdg"):
                has_pdg = True
                break
        if has_pdg:
            pdg_folders.append(temp_folder)
        else:
            for f in os.scandir(temp_folder):
                if f.is_dir():
                    temp_folders.append(f.path)
    return pdg_folders


def set_task(app, dlg, pdg_folder):
    # 打开文件选择框
    dlg.children()[3].click()
    select_dlg = app.window(title=u'选择存放PDG文件的文件夹')
    select_dlg.children()[6].set_text(pdg_folder)
    ok_btn = select_dlg.children()[9]
    win32gui.SetWindowLong(ok_btn.handle, win32con.GWL_STYLE,
                           win32gui.GetWindowLong(ok_btn.handle, win32con.GWL_STYLE) & ~win32con.WS_DISABLED)
    ok_btn.click()

    while True:
        app.window(title=u'格式统计').wait('ready', timeout=9999).children()[0].click()
        dlg.children()[0].click()
        if not app.window(title=u'格式统计').exists():
            break


pdg_folders = get_pdg_folders(请将总文件夹路径填到这里，将递归查找文件夹下所有含pdg文件的子文件夹)
total = len(pdg_folders)
print(f'待转换总数: {total}')

apps = [Application(backend='win32').start('./Pdg2Pic.exe') for _ in range(N)]
states = [0 for _ in range(N)]  # 0: 闲置  1: 正在转换
current = 0

while len(pdg_folders) + sum(states) != 0:
    app = apps[current]
    dlg = apps[current].Pdg2Pic
    if states[current] == 0:  # 闲置
        if len(pdg_folders) == 0:  # 没有剩余任务了
            current = (current + 1) % N
            continue
        else:  # 有剩余任务
            pdg_folder = pdg_folders.pop()
            set_task(app, dlg, pdg_folder)
            states[current] = 1
            print(f'进度: {total - len(pdg_folders)} / {total}')
    elif states[current] == 1:  # 正在转换
        if len(app.windows()) == 11:  # 转换完成
            while True:
                app.window(
                    title='Pdg2Pic', predicate_func=lambda dlg: len(dlg.children()) == 3
                ).wait('ready', timeout=9999).children()[0].click()
                if not app.window(
                    title='Pdg2Pic', predicate_func=lambda dlg: len(dlg.children()) == 3
                ).exists():
                    break
            if len(pdg_folders) == 0:  # 没有剩余任务了
                states[current] = 0
            else:  # 还有任务，直接分配
                pdg_folder = pdg_folders.pop()
                set_task(app, dlg, pdg_folder)
                print(f'进度: {total - len(pdg_folders)} / {total}')
    current = (current + 1) % N
    time.sleep(0.05)

for app in apps:
    app.kill()

print('转换完成！')
print('所有pdf路径如下：')
pdg_folders = get_pdg_folders('C:\\Users\\86188\\Desktop\\download_file')
pdf_files = [f + '.pdf' for f in pdg_folders]
ok_pdf_files = []
for pdf_file in pdf_files:
    if os.path.isfile(pdf_file):
        ok_pdf_files.append(pdf_file)
print(f"成功转换数：{len(ok_pdf_files)}")
print('所有pdf路径：')
print(ok_pdf_files)
