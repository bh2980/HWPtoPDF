from glob import glob
import win32com.client as win32
import tkinter as tk
import tkinter.filedialog as fd
from os import path

init_dir = path.join(path.expanduser('~'),'Desktop')

def pathload(arrow, textfield):
    global init_dir
    init_dir = fd.askdirectory(initialdir=init_dir, title="Select %s folder" % arrow)

    textfield.delete(0, 'end')
    textfield.insert(0, init_dir)

def hwpToPdf(file, path):
    if file == '' or path == '':
        print("[에러] 경로가 지정되지 않았습니다.")

        return
    else:
        print("-------------------------필독!-------------------------\n")
        print("\"한글을 이용하여 위 파일에 접근하려는 시도가 있습니다.\"")
        print("와 같은 팝업창이 뜰 경우 [모두 허용]을 눌러주세요.\n")
        print("-------------------------------------------------------")
        file = glob(file + '\*.hwp')

        hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

        for i in file:
            hwp.Open(i)
            i = i.split('\\')
            i.reverse()
            hwp.SaveAs(path + '/' + i[0].replace('.hwp', '.pdf'), "PDF")
            print("변환 완료 :", path + '/' + i[0])

        hwp.Quit()

        return

root = tk.Tk()
root.title("HWPDF")
root.geometry("250x220")
root.resizable(False, False)

hwpfolder = tk.Label(root, text="HWP 파일 폴더")
hwpfolder.place(x=30, y=30)

filedirtext = tk.Entry(root)
filedirtext.place(x=30, y=50, width=135, height=25)

filebtn = tk.Button(root, text="Load", command=lambda: pathload("hwp file", filedirtext))
filebtn.place(x=170, y=50, width=50)

savefolder = tk.Label(root, text="저장 폴더")
savefolder.place(x=30, y=90)

pathdirtext = tk.Entry(root)
pathdirtext.place(x=30, y=110, width=135, height=25)

pathbtn = tk.Button(root, text="Load", command=lambda: pathload("save", pathdirtext))
pathbtn.place(x=170, y=110, width=50)

runbtn = tk.Button(root, text="Run!", height=3, command=lambda: hwpToPdf(filedirtext.get(), pathdirtext.get()))
runbtn.pack(side="bottom", fill="x")

root.mainloop()