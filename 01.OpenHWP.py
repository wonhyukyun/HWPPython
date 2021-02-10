import win32com.client as win32
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True

act = hwp.CreateAction("PageHiding")
set = act.CreateSet()
act.GetDefault(set)
set.SetItem("Fields", 32)
act.Execute(set)
print(set.Item("Fields"))
