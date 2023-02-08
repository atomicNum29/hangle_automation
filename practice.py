import win32com.client as win32

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True

hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "ㅁㄴㅇㄹ"
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.InsertPicture("D:\\hangle_automation\\data\\스크린샷(1).png", Embedded=True) # 이미지 삽입