#https://www.martinii.fun/102?category=766606
"""
2번 질문======================
제 1조(상호)
제 2조(목적)을 다음과 같이 굵게(BOLD) 하고 싶습니다.
제 1조(상호) 굵게
제 2조(목적) 굵게
이와 같이 굵게 칠하고 싶습니다. -
"""
 
 
from time import sleep  # 너무 빨리 실행돼서 오류가 났나 싶어 써본 기능(일시정지)
from tkinter import Tk  # 다이얼로그창 모듈 임포트
from tkinter.filedialog import askopenfilename  # HWP파일선택을 위한 다이얼로그창
import re  # 정규표현식
import win32com.client as win32  # 한/글 오브젝트 조작을 위한 win32 임포트
 
 
def hwp_get_max_number(hwp):  # 조항 갯수 추출코드
	"""중간에 탐색을 멈추는 현상이 지속적으로 발생하여, 
    마지막 조항에 도착하지 않았으면 계속 탐색을 처음부터 재시도하도록 코드 변경."""
    text = hwp.GetTextFile("TEXT", "").split("\r\n")  # 모든 텍스트를 전부 긁어서
    max_number = max([int(i.replace(" ", "")[1:].split("조")[0].strip()) for i in text if i.startswith("제")])
    # "제?조"로 시작하는 문자열의 ?만 추출해서 정수로 바꾸고 그 리스트 중 최대값을 max_number로 지정함
    return max_number
 
 
def hwp_get_bold(hwp, type):  # "굵게" 적용 전에 현재 굵은 상태인지 체크하는 코드 추가(위6줄)
    Act = hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
    Set = Act.CreateSet()  # 세트 생성
    Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
    if Set.Item("Bold") == 1:  # 파라미터셋테이블에서 "CharShape" 검색, 아이템아이디에서 "Bold" 찾음
        return  # 굵은 상태면 함수종료. 
    else:  # 0이나 None을 리턴하는 경우가 있는데, 0은 모두 진하지 않은 경우, None은 섞인 경우. 둘 다 아래코드 실행됨
        if type == "장":  # "제?장"의 경우
            hwp.HAction.Run("MoveSelLineEnd")  # 라인 끝까지 선택하고
            hwp.HAction.Run("CharShapeBold")  # 굵게Ctrl-B 커맨드 입력
            hwp.HAction.Run("MoveLineBegin")  # 다시 라인 시작점으로 이동
        elif type == "조":  # "제?조"의 경우
            # 아래는 라인 시작점에서 가장 먼저 나오는 ")"를 찾아서 그 지점부터 시작점까지 선택하고 진하게 적용.
            hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HParameterSet.HFindReplace.FindString = ")"
            hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
            hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            hwp.HParameterSet.HFindReplace.FindType = 1
            hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HAction.Run("MoveRight")
            hwp.HAction.Run("MoveLeft")
            hwp.HAction.Run("MoveSelLineBegin")
            hwp.HAction.Run("CharShapeBold")
            hwp.HAction.Run("MoveLineBegin")
        else:  # 장도 조도 아닌 게 선택되어 있으면?
            raise TypeError  # 선택에서든 탐색에서든 명백한 오류이므로 에러발생
 
 
def hwp_init(filename):  # 아래아한글 열기
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 오브젝트 생성
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈 실행
    hwp.Open(filename)  # 최초에 선택한 HWP파일 불러오기
    hwp.XHwpWindows.Item(0).Visible = True  # 숨김해제(열 때 기본값이 "숨김")
    hwp.HAction.Run("FrameFullScreen")  # 전체화면
    return hwp
 
 
def hwp_find_and_get_bold(hwp):  # 찾아가서 굵게 하는 메인함수.
    max_number = hwp_get_max_number(hwp)  # 우선 마지막 조항번호를 알아내고
    hwp.InitScan() # 탐색시작 커맨드
    장번호 = 1  # 장-인덱스
    조항번호 = 1  # 조-인덱스
    while True:  # break 만날 때까지 무한반복
        text = hwp.GetText()  # 다음라인 탐색
        if text[0] == 1 and 조항번호 != max_number:  # 검색이 갑자기 종료되는 버그(문서 손상? 메모리 문제? 초기화하면 정상작동함..)
            hwp.ReleaseScan()  # 마지막 조항까지 도착 안했는데 탐색이 종료되면, 탐색프로세스 종료하고
            hwp.InitScan()  # 새로운 탐색프로세스 생성
            sleep(0.1)  # 쉬어가자. 
            continue  # while문 재개
        elif 조항번호 > max_number:  # 마지막 조항까지 도달했으면
            break  # while문 종료
        else:  # 그 외 일반적인 경우에는 
            if re.match(rf"^제{조항번호}조\(?", text[1].replace(" ", "")):  # "제?조"를 찾았으면
                hwp.MovePos(201)  # 해당 라인 시작점으로 가서
                hwp_get_bold(hwp, "조") # 위의 굵게함수 적용
                조항번호 += 1  # 조-인덱스 + 1
                hwp.InitScan()  # 탐색 재시작(문자열 수정하면 탐색 재시작해야 함)
            if re.match(rf"^제{장번호}장", text[1].replace(" ", "")):  # "제?장"을 찾았으면
                hwp.MovePos(201)  # 해당 라인 시작점으로 가서
                hwp_get_bold(hwp, "장")  # 위의 굵게함수 적용
                장번호 += 1  # 장-인덱스 + 1
                hwp.InitScan()  # 탐색 재시작
            else:  # 장도 절도 아니면
                pass  # 아무 것도 하지 말고 넘어가기
    hwp.ReleaseScan()  # while문 종료되었으면 탐색프로세스도 종료
    hwp.MovePos(2)  # 끝났으니까 문서 시작점으로 이동.  끝.
 
 
if __name__ == '__main__':  # 메인함수(필수는 아닌데, 뭔가 진입점 같은 게 간결해 보임?)
    root = Tk()  # GUI클래스 생성
    filename = askopenfilename() # HWP파일선택 다이얼로그 열기
    root.destroy()  # 파일 선택했으면 GUI 닫기
 
    hwp = hwp_init(filename=filename)  # 한/글 오브젝트 생성
    hwp.MovePos(2)  # 문서 시작점으로 이동한 후에
    hwp_find_and_get_bold(hwp)  # 위의 찾아가서 진하게 하는 함수 실행. 끝.