from glob import glob
import win32com.client as win32

file = glob(r"C:\Users\bh2980\Desktop\test\*.hwp")

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

for i in file:
    hwp.Open(i)
    hwp.SaveAs(i.replace('.hwp', '.pdf'), "PDF")

hwp.Quit()