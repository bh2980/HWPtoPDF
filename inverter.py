from glob import glob
import win32com.client as win32

file = glob(r"경로")

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

for i in file:
    hwp.Open(i)
    hwp.SaveAs(i.replace('.hwp', '.pdf'), "PDF")

hwp.Quit()
