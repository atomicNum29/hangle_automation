import win32com.client as win32
import os

data_path = "D:/hangle_automation/data"
original_hwpfile_name = "(매뉴얼) 운영_매뉴얼_ver1_최종 (1) - 복사본.hwp"
image_file_path = "D:/hangle_automation/data/juniversity 출강증빙 1.25_2.3/학부모"
# image_file_name = "01.26/*.png"

output_path = "D:/hangle_automation/output"
output_file_name = "학부모_"

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

for date in ['01.26', '01.27', '01.28', '01.30', '01.31', '02.01', '02.02']:
    # 특정 한글 파일 열기
    hwp.Open(data_path + '/' + original_hwpfile_name,"HWP","forceopen:true")
    for image_file_name in os.listdir(image_file_path + '/' + date):
        # 이미지 삽입할 위치 찾기; image1 위치를 찾는다.
        hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
        hwp.HParameterSet.HFindReplace.FindString = "image"+ image_file_name.removesuffix('.jpg')
        hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
        # 이미지 삽입
        hwp.InsertPicture(image_file_path + '/' + date + '/' + image_file_name, Embedded=True)
        # 붙여넣은 이미지 속성 변경으로 크기 조정
        hwp.FindCtrl()
        hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
        hwp.HParameterSet.HShapeObject.Height = 10000
        hwp.HParameterSet.HShapeObject.HeightRelTo = hwp.HeightRel("Page")
        hwp.HParameterSet.HShapeObject.Width = 10000
        hwp.HParameterSet.HShapeObject.WidthRelTo = hwp.WidthRel("Page")
        hwp.HParameterSet.HShapeObject.HorzRelTo = hwp.HorzRel("Para")
        hwp.HParameterSet.HShapeObject.HSet.SetItem("ShapeType", 1)
        hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
    
    # 수정 완료한 파일 저장 
    hwp.SaveAs(output_path + '/' + output_file_name + date + ".hwp")

hwp.Quit()