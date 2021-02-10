"""
3번 질문==============
정관
제 1 장 총칙
제 1조
제 2조..
제 3조...
제 2 장 주식과 주권
제 4조....
제 5조..
제 3 장 임원
제 6조
등과 같을 때.
 
장의 위치를 페이지 가운데로 위치하고 싶고, 또한
각 장 줄의 위와 아래에 빈 줄을 삽입해 넣고 싶습니다. 즉,
 
                        빈줄
                        정관
                        빈줄
                    제 1 장 총칙
                        빈줄
제 1조
제 2조..
제 3조...
                        빈줄
                제 2 장 주식과 주권
                        빈줄
제 4조....
제 5조..
                        빈줄
                    제 3 장 임원
                        빈줄
제 6조
 
등과 같이 만들고 싶습니다.
4번 질문==============
제 1조...
제 2조....
제 3조...
라고 했을 때
각 조 사이 마다 빈 줄을 하나씩 삽입하고 싶습니다 .즉
제 1조...
빈줄
제 2조....
빈줄
제 3조...
빈줄
이와 같이 만들고 싶습니다.
이것을 이미 만들어 주신 조항번호 파이썬 소스에 제가 부탁드린 부분만을 수정하여서
4가지로 파이썬 파일로 각각 만들어 주신다면 그것을 바탕으로 눈이 빠져라 열심히 공부를 하겠음을 약속 드리겠습니다.
좀 빠르게 배우고 싶은 마음에.
만들어진 소스를 바탕으로 역으로 공부해 가는 방향을 선택하고 싶습니다.
정말 어처구니 없는 부탁 같아서. 염치가 없습니다.
그래도 제 마음속에 염원을 말씀 드려 보았습니다.
죄송하고 부끄럽습니다....
 
 
 => "제1장"이나 "제 1 장"을 모두 찾아내는 경우는 정규식을 쓰면 참 간단하지만, (X)
    정규식 없이 str.replace(" ","") 방식으로 문자열에서 스페이스를 제거하고 검색하는 방법도 있다. (ㅇ)
    "제"와 "장"으로 쪼갠 0번 인덱스 원소가 정수면 된다.
 
    빈 줄을 삽입하는 코드는 자칫 반복 실행하면 두 줄, 세 줄로 늘어나버릴 수도 있으므로,
    부담스럽지만 빈 줄을 모두 제거한 후에(제목 위아래, 장 위아래 말고는 빈 줄이 없을 것이므로(?))  # "^n^n" -> "^n"
    빈줄삽입 코드를 실행한다. (X)
    혹은, 찾은 문자열 위가 빈줄인지, 아래도 빈줄인지 검색만 하는 코드로도 대체가 가능하다. (ㅇ)
 
    마지막으로, 위는 빈 줄이 있는데 아래에만 빈 줄이 없는 경우가 있으므로 둘을 따로 점검한다.
 
 => 4번 질문은 간단하다. 조제목 위에 빈줄이 있는지 확인하고 없으면 엔터. 아래는 확인할 필요가 없다.
"""
 
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re
import win32com.client as win32
import pyperclip as cb
 
 
def hwp_init(filename):
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.Open(filename)
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.HAction.Run("FrameFullScreen")
    return hwp
 
 
def hwp_center_align_and_insert_blank_line(hwp, dir, target):
    if target == "장":
        hwp.HAction.Run("ParagraphShapeAlignCenter")
    else:
        pass
 
    if dir == "above":
        hwp.HAction.Run("MoveLineBegin")
        hwp.HAction.Run("BreakPara")
    elif dir == "below":
        hwp.HAction.Run("MoveLineEnd")
        hwp.HAction.Run("BreakPara")
    else:
        raise ValueError
 
 
def hwp_check_if_blank_exists_above(hwp):
    current_position = hwp.GetPos()  # 현위치 저장(간혹 다음 검색위치로 튀는문제 조치)
    hwp.HAction.Run("MoveLineBegin")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("Copy")
    hwp.SetPos(*current_position)  # 방금위치 복원
    if cb.paste() == "\r\n\r\n":
        return True
    else:
        return False
 
 
def hwp_check_if_blank_exists_below(hwp):
    current_position = hwp.GetPos()  # 현위치 저장(간혹 다음 검색위치로 튀는문제 조치)
    hwp.HAction.Run("MoveLineEnd")
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("Copy")
    hwp.SetPos(*current_position)  # 방금위치 복원
    if cb.paste() == "\r\n\r\n":
        return True
    else:
        return False
 
 
def hwp_find_and_go(hwp):
    hwp.InitScan()
    장번호 = 1
    조번호 = 1
    while True:
        text = hwp.GetText()
        if text[0] == 1:
            break
        else:
            if re.match(rf"^제{장번호}장.+", text[1].strip().replace(" ", "")):
                장번호 += 1
                hwp.MovePos(201)  # moveScanPos : GetText() 실행 후 위치로 이동한다.
                if not hwp_check_if_blank_exists_above(hwp):
                    hwp_center_align_and_insert_blank_line(hwp, "above", "장")
                if not hwp_check_if_blank_exists_below(hwp):
                    hwp_center_align_and_insert_blank_line(hwp, "below", "장")
                hwp.InitScan()
 
            if re.match(rf"^제{조번호}조.+", text[1].strip().replace(" ", "")):
                조번호 += 1
                hwp.MovePos(201)  # moveScanPos : GetText() 실행 후 위치로 이동한다.
                if not hwp_check_if_blank_exists_above(hwp):
                    hwp_center_align_and_insert_blank_line(hwp, "above", "조")
                    hwp.InitScan()
 
            else:
                pass
    hwp.ReleaseScan()
    hwp.MovePos(2)
 
 
if __name__ == '__main__':
    root = Tk()
    filename = askopenfilename()
    root.destroy()
    hwp = hwp_init(filename=filename)
    hwp_find_and_go(hwp)