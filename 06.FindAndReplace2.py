#https://www.martinii.fun/99?category=766606
from tkinter import Tk  # GUI 띄우는 창
from tkinter.filedialog import askopenfilename  # HWP파일 선택하기 위한 다이얼로그창
import re  # 정규표현식
import win32com.client as win32  # 한/글 열기 위한 모듈
 
 
def hwp_find_replace(find_string, replace_string):  # 한/글 찾아바꾸기 함수(녹화한 스크립트임)
    hwp.Run("MoveSelNextWord")
    hwp.HAction.GetDefault("ExecReplace", hwp.HParameterSet.HFindReplace.HSet)  # 한/글 특성상 부득이하게 두두번번 실실행행
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
    hwp.HParameterSet.HFindReplace.FindString = find_string
    hwp.HParameterSet.HFindReplace.ReplaceString = replace_string
    hwp.HParameterSet.HFindReplace.ReplaceMode = 1
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HAction.Execute("ExecReplace", hwp.HParameterSet.HFindReplace.HSet)  # 이 시점에 찾기만 하고 바뀌지 않음
    hwp.HAction.GetDefault("ExecReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
    hwp.HParameterSet.HFindReplace.FindString = find_string
    hwp.HParameterSet.HFindReplace.ReplaceString = replace_string
    hwp.HParameterSet.HFindReplace.ReplaceMode = 1
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HAction.Execute("ExecReplace", hwp.HParameterSet.HFindReplace.HSet)  # 이 시점에 변경 완료
    hwp.Run("Cancel")
 
 
def hwp_init(filename):  # 한/글 여는 코드
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 아래아한글 열고
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈 불러오고(파일 열거나 저장, 이미지 불러올 때 보안팝업이 뜨지 않음)
    hwp.Open(filename)  # 해당문서 열기
    hwp.XHwpWindows.Item(0).Visible = True  # 숨김해제(최신버전 기준 백그라운드에서 시작함)
    hwp.HAction.Run("FrameFullScreen")  # 전체화면
    return hwp  # 생성한 한/글 오브젝트를 리턴
 
 
def hwp_reindex(hwp):  # 번호 재정렬하는 메인함수
    hwp.InitScan()  # 검색준비
    조항번호 = 1  # 인덱스
    while True:  # break 만나기 전까지 무한반복~
        text = hwp.GetText()  # 문서를 엔터로 쪼갠 리스트 탐색
        if text[0] == 1:  # 종료 코드인 1이 발생하면
            break  # while문도 종료
        else:  # 그 전까지는
            if re.match(r"^제\d+조\(?", text[1]) and text[1].startswith(f"제{조항번호}조("):
                # 문단이 "제?조("로 시작하면서, 조항번호가 올바르게 들어가 있는 경우(조항순서==인덱스)
                조항번호 += 1  # 인덱스만 하나 올리고 넘어감(문서는 바뀌지 않음)
                continue
            elif re.match(r"^제\d+조\(?", text[1]) and not text[1].startswith(f"제{조항번호}조("):
                # 문단이 "제?조("로 시작하는데, 조항번호가 올바르지 않은 경우(조항순서가 인덱스와 다름)
                hwp.MovePos(201)  # moveScanPos : GetText() 로 탐색중인 현재 위치로 이동
                hwp.HAction.Run("MoveLineBegin")  # 해당라인 앞으로 이동(이 라인은 없어도 무관하나 모니터링을 위해 추가함. 
                # MovePos(201)만 실행하면 해당 캐럿으로 이동하지 않음)
                hwp_find_replace(re.match(r"^제\d+조\(?", text[1]).group(0), f"제{조항번호}조(")
                # 위에서 정의한 찾아바꾸기 함수로 "제?조(" 안의 ?를 올바른 번호로 대체함
                조항번호 = 1  # 조항번호를 1로 바꾸고
                hwp.InitScan()  # 검색을 다시 실행함(왜? 문자열 변동이 생기면 검색 강제종료됨)
            else:
                pass  # 위 탐색과정을 반복함
    hwp.ReleaseScan()  # break문을 만난 후에는 "검색종료" 코드를 실행하고
    hwp.MovePos(2)  # 완료되었으니 문서 맨 위로 이동
 
 
if __name__ == '__main__':  # 메인 파트(실행부분)
    root = Tk()  # 내장 GUI모듈 불러와서
    filename = askopenfilename()  # 파일선택창 열고, 선택한 파일명을 filename에 지정
    root.destroy()  # 파일선택창 종료
    hwp = hwp_init(filename=filename)  # 아래아한글 시작하면서 선택한 파일 열기
    hwp_reindex(hwp)  # 메인함수 실행