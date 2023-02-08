import win32com.client as win32

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
hwp.XHwpWindows.Item(0).Visible = True

def printhwp(s):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = s
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

hwp.Open("D:\hangle_automation\data\(매뉴얼) 운영_매뉴얼_ver1_최종 (1) - 복사본.hwp","HWP","forceopen:true")

hwp.InsertPicture("D:\hangle_automation\data\스크린샷(1).png", Embedded=True) # 이미지 삽입
hwp.FindCtrl() # 이미지 선택 (커서에서 인접한 개체 선택)
hwp.HAction.Run("Cut") # 잘라내기
while True:
	# 이미지 삽입할 위치 찾기
	hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
	hwp.HParameterSet.HFindReplace.FindString = "image1"
	hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
	result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

	# 다 바꿨으면 종료
	if result == False:
		break

	# 붙여넣기
	hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
	hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)

hwp.FindCtrl()
hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
hwp.HParameterSet.HShapeObject.Height = 17007
hwp.HParameterSet.HShapeObject.Width = 22677
# hwp.HParameterSet.HShapeObject.HorzRelTo = hwp.HorzRel("Para")
# hwp.HParameterSet.HShapeObject.HSet.SetItem("ShapeType", 1)
print(hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet))

# hwp.Save