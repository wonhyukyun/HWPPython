import win32com.client as win32
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True

act = hwp.Open("C:\\Users\\wonhyuk.yun\\Desktop\\test01.hwp")

current_Page = hwp.XHwpDocuments.Item(0).XHwpDocumentInfo.CurrentPage
print(current_Page)

