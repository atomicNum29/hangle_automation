import win32com.client as win32

data_path = "D:/hangle_automation/data"
original_hwpfile_name = "(매뉴얼) 운영_매뉴얼_ver1_최종 (1) - 복사본.hwp"
image_file_name = "스크린샷(1).png"

output_path = "D:/hangle_automation/output"
output_file_name = "저장.hwp"

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
# 보안 창 생략
hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
# 한글 프로그램 보여주기
hwp.XHwpWindows.Item(0).Visible = True

# 현재 커서 위치에 s 문자열 작성
def printhwp(s):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = s
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

# 특정 한글 파일 열기
hwp.Open(data_path + '/' + original_hwpfile_name,"HWP","forceopen:true")
# 특정 이미지 삽입 후 잘라내기
# 클립보드에 이미지를 저장함.
hwp.InsertPicture(data_path + '/' + image_file_name, Embedded=True) # 이미지 삽입
hwp.FindCtrl() # 이미지 선택 (커서에서 인접한 개체 선택)
hwp.HAction.Run("Cut") # 잘라내기
# 이미지 삽입할 위치 찾기; image1 위치를 찾는다.
hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
hwp.HParameterSet.HFindReplace.FindString = "image1"
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
# 붙여넣기
hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
# 붙여넣은 이미지 속성 변경으로 크기 조정
hwp.FindCtrl()
hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
hwp.HParameterSet.HShapeObject.Height = 17007
hwp.HParameterSet.HShapeObject.Width = 22677
hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
 
# 수정 완료한 파일 저장 
hwp.SaveAs(output_path + '/' + output_file_name)
hwp.Quit()