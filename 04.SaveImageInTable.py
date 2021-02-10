#https://www.martinii.fun/96?category=766606
# %% 임포트
 
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import os
import win32com.client as win32
 
# %% 이미지파일 선택
 
root = Tk()  # 이미지선택창 열기
imagelist = askopenfilenames()
root.destroy()  # 이미지선택창 닫기
 
BASE_DIR = imagelist[0].rsplit("/", maxsplit=1)[0]  # 이미지리스트에서 경로 추출
imagelist = [i.rsplit("/", maxsplit=1)[1] for i in imagelist]  # 이미지리스트에서 파일명만 남김
 
# %% 표_리스트 만들기
 
표_리스트 = list(set([i.split("_")[0][:-1] for i in imagelist]))
표_리스트.sort()
 
# %% 한/글 오브젝트 생성
 
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 오브젝트 생성
hwp.XHwpWindows.Item(0).Visible = True  # 숨김해제
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 실행하는 컴퓨터에서 레지스트리 등록해야 함
# 참고영상 : https://www.martinii.fun/entry/파이썬-아래아한글-보안모듈-설치방법귀찮은-보안팝업-제거-1
 
# %% 여백조정
 
hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)
hwp.HParameterSet.HSecDef.PageDef.Landscape = 1  # 가로로
hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(25.0)
hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(20.0)
hwp.HParameterSet.HSecDef.PageDef.TopMargin = hwp.MiliToHwpUnit(15.0)
hwp.HParameterSet.HSecDef.PageDef.BottomMargin = hwp.MiliToHwpUnit(15.0)
hwp.HParameterSet.HSecDef.PageDef.HeaderLen = hwp.MiliToHwpUnit(0.0)
hwp.HParameterSet.HSecDef.PageDef.FooterLen = hwp.MiliToHwpUnit(0.0)
hwp.HParameterSet.HSecDef.PageDef.GutterLen = hwp.MiliToHwpUnit(0.0)
hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)
hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)  # 문서 전체 변경
hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)
 
for idx, content in enumerate(표_리스트):
    # %% 표 생성
    
    hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = 5
    hwp.HParameterSet.HTableCreation.Cols = 3
    hwp.HParameterSet.HTableCreation.WidthType = 2
    hwp.HParameterSet.HTableCreation.HeightType = 1
    hwp.HParameterSet.HTableCreation.WidthValue = hwp.MiliToHwpUnit(250)  # 표만들기 포스팅 참고 
    hwp.HParameterSet.HTableCreation.HeightValue = hwp.MiliToHwpUnit(205)
    hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 5)
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, hwp.MiliToHwpUnit(79.73))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, hwp.MiliToHwpUnit(79.73))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, hwp.MiliToHwpUnit(79.73))
    hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 5)
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(0, hwp.MiliToHwpUnit(10.0))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(1, hwp.MiliToHwpUnit(10.0))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(2, hwp.MiliToHwpUnit(50))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(3, hwp.MiliToHwpUnit(50))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(4, hwp.MiliToHwpUnit(50))
    hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1  # 글자처럼 취급
    hwp.HParameterSet.HTableCreation.TableProperties.Width = hwp.MiliToHwpUnit(250)
    hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
 
    # %% 라인 투명 설정
    
    hwp.HAction.Run("TableCellBlockRow")  # 표 1행 선택
    hwp.HAction.GetDefault("CellBorder", hwp.HParameterSet.HCellBorderFill.HSet)
    hwp.HParameterSet.HCellBorderFill.TypeVert = hwp.HwpLineType("None")  # 세로줄 투명
    hwp.HAction.Execute("CellBorder", hwp.HParameterSet.HCellBorderFill.HSet)
 
    hwp.HAction.Run("TableColPageDown")  # 표 전체 선택
    hwp.HAction.GetDefault("CellBorder", hwp.HParameterSet.HCellBorderFill.HSet)
    hwp.HParameterSet.HCellBorderFill.BorderTypeTop = hwp.HwpLineType("None")  # 상단 투명
    hwp.HParameterSet.HCellBorderFill.BorderTypeRight = hwp.HwpLineType("None")  # 우측 투명
    hwp.HParameterSet.HCellBorderFill.BorderTypeLeft = hwp.HwpLineType("None")  # 좌측 투명
    hwp.HAction.Execute("CellBorder", hwp.HParameterSet.HCellBorderFill.HSet)
    hwp.HAction.Run("Cancel")  # 셀선택 해제
 
    # %% 커서를 2행(A2)으로 이동
    
    # hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("TableColBegin")
    hwp.HAction.Run("TableColPageUp")
    hwp.HAction.Run("TableLowerCell")
    hwp.HAction.Run("Cancel")
 
    # %% 제목 3개 셀 색 채우기
    
    hwp.HAction.GetDefault("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
    hwp.HParameterSet.HCellBorderFill.FillAttr.type = hwp.BrushType("NullBrush|WinBrush")
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceColor = hwp.RGBColor(66, 199, 241)
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushHatchColor = hwp.RGBColor(153, 153, 153)
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceStyle = hwp.HatchStyle("None")
    hwp.HParameterSet.HCellBorderFill.FillAttr.WindowsBrush = 1
    hwp.HAction.Execute("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
 
    hwp.HAction.Run("TableRightCell")  # 우측셀로 이동
 
    hwp.HAction.GetDefault("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
    hwp.HParameterSet.HCellBorderFill.FillAttr.type = hwp.BrushType("NullBrush|WinBrush")
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceColor = hwp.RGBColor(50, 215, 9)
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushHatchColor = hwp.RGBColor(153, 153, 153)
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceStyle = hwp.HatchStyle("None")
    hwp.HParameterSet.HCellBorderFill.FillAttr.WindowsBrush = 1
    hwp.HAction.Execute("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
 
    hwp.HAction.Run("TableRightCell")  # 우측셀로 이동
 
    hwp.HAction.GetDefault("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
    hwp.HParameterSet.HCellBorderFill.FillAttr.type = hwp.BrushType("NullBrush|WinBrush")
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceColor = hwp.RGBColor(209, 209, 209)
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushHatchColor = hwp.RGBColor(153, 153, 153)
    hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceStyle = hwp.HatchStyle("None")
    hwp.HParameterSet.HCellBorderFill.FillAttr.WindowsBrush = 1
    hwp.HAction.Execute("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
 
    # %% 제목행 다시 선택
 
    hwp.HAction.Run("TableColBegin")
    hwp.HAction.Run("TableColPageUp")
    hwp.HAction.Run("TableCellBlockRow")
 
    # %% 휴먼명조, 16pt로 변경
 
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.FaceNameUser = "휴먼명조"
    hwp.HParameterSet.HCharShape.FontTypeUser = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameSymbol = "휴먼명조"
    hwp.HParameterSet.HCharShape.FontTypeSymbol = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameOther = "휴먼명조"
    hwp.HParameterSet.HCharShape.FontTypeOther = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameJapanese = "휴먼명조"
    hwp.HParameterSet.HCharShape.FontTypeJapanese = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameHanja = "휴먼명조"
    hwp.HParameterSet.HCharShape.FontTypeHanja = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.SizeLatin = 100
    hwp.HParameterSet.HCharShape.FaceNameLatin = "휴먼명조"
    hwp.HParameterSet.HCharShape.FontTypeLatin = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.SizeHangul = 100
    hwp.HParameterSet.HCharShape.FaceNameHangul = "휴먼명조"
    hwp.HParameterSet.HCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.RatioUser = 100
    hwp.HParameterSet.HCharShape.SizeUser = 100
    hwp.HParameterSet.HCharShape.RatioSymbol = 100
    hwp.HParameterSet.HCharShape.SizeSymbol = 100
    hwp.HParameterSet.HCharShape.RatioOther = 100
    hwp.HParameterSet.HCharShape.SizeOther = 100
    hwp.HParameterSet.HCharShape.SpacingJapanese = 0
    hwp.HParameterSet.HCharShape.RatioJapanese = 100
    hwp.HParameterSet.HCharShape.SizeJapanese = 100
    hwp.HParameterSet.HCharShape.SpacingHanja = 0
    hwp.HParameterSet.HCharShape.RatioHanja = 100
    hwp.HParameterSet.HCharShape.SizeHanja = 100
    hwp.HParameterSet.HCharShape.SpacingLatin = 0
    hwp.HParameterSet.HCharShape.RatioLatin = 100
    hwp.HParameterSet.HCharShape.SpacingHangul = 0
    hwp.HParameterSet.HCharShape.RatioHangul = 100
    hwp.HParameterSet.HCharShape.OffsetUser = 0
    hwp.HParameterSet.HCharShape.SpacingUser = 0
    hwp.HParameterSet.HCharShape.OffsetSymbol = 0
    hwp.HParameterSet.HCharShape.SpacingSymbol = 0
    hwp.HParameterSet.HCharShape.OffsetOther = 0
    hwp.HParameterSet.HCharShape.SpacingOther = 0
    hwp.HParameterSet.HCharShape.OffsetJapanese = 0
    hwp.HParameterSet.HCharShape.OffsetHanja = 0
    hwp.HParameterSet.HCharShape.OffsetLatin = 0
    hwp.HParameterSet.HCharShape.OffsetHangul = 0
    hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(16.0)
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
 
    # %% 제목행에 문자열 삽입
 
    hwp.HAction.Run("TableColBegin")
 
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "장소명 : "
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.CreateField(Direction="장소명", memo="장소명을 입력하세요.", name=f"{idx}_name")
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "코드명 : "
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.CreateField(Direction="장소코드", memo="장소 코드명을 입력하세요.", name=f"{idx}_code")
    hwp.HAction.Run("ParagraphShapeAlignCenter")
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "타입 : "
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.CreateField(Direction="타입", memo="자료타입을 입력하세요.", name=f"{idx}_type")
    hwp.HAction.Run("ParagraphShapeAlignCenter")
 
    # %% 아래 열 글자크기 12로
    
    hwp.HAction.Run("TableRightCell")  # 2행으로 이동
    hwp.HAction.Run("TableCellBlockRow")  # 2행 전체 선택
    hwp.HAction.Run("CharShapeHeightIncrease")  # 폰트+=1
    hwp.HAction.Run("CharShapeHeightIncrease")  # 폰트+=1
    hwp.Run("TableColBegin")  # 2행1열로 이동
 
    # %% 텍스트 삽입
    
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "시간-폭"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("ParagraphShapeAlignCenter")
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "주파수-폭"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("ParagraphShapeAlignCenter")
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "주파수-시간-폭"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("ParagraphShapeAlignCenter")
 
    # %% 이미지 삽입할 9개의 셀에 각각 필드명 삽입(나중에 편함)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_E_time"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_E_PDF"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_E_temporal"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_N_time"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_N_PDF"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_N_temporal"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_Z_time"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_Z_PDF"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    hwp.HAction.Run("TableRightCell")
 
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}_Z_temporal"
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
 
    # %% 텍스트 필드에 문자열 삽입
    
    hwp.PutFieldText(f"{idx}_name", content.split("..")[0])
    hwp.PutFieldText(f"{idx}_code", content.split(".")[1])
    hwp.PutFieldText(f"{idx}_type", "(속도" + content.split("..")[-1] + "*)")
 
    # %% 이미지 삽입
    
    slot_list = [f"{idx}_E_time", f"{idx}_E_PDF", f"{idx}_E_temporal", f"{idx}_N_time", f"{idx}_N_PDF",
                 f"{idx}_N_temporal", f"{idx}_Z_time", f"{idx}_Z_PDF", f"{idx}_Z_temporal"]  # 이미지 삽입할 셀 필드 리스트
    for j in slot_list:
        hwp.MoveToField(j)  # 해당 필드로 이동
        hwp.InsertPicture(os.path.join(BASE_DIR, f"{content + j.split('_', maxsplit=1)[1]}.png"), Embedded=True,
                          sizeoption=3)  # 이미지 삽입
    hwp.Run("MoveDocEnd")  # 문서 끝으로 이동
    hwp.Run("DeleteBack")  # 엔터 하나 삭제
 
########### hwp.Save는 생략함. 끝.