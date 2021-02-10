#https://www.martinii.fun/100?category=766606
"""
1번 질문:===========================
제[공백][공백]1조
제[공백][공백]2조
.
.
제[공백]13조
제[공백]14조
등과 같이 앞에 3자리를 기준으로 빈 공백을 놓아 두고 싶습니다....
그러면 ”조“자의 위치가 나란히 위치 할 수 있어서 그럽니다.
 
 => 조항번호를 매길 때
    숫자가 한 자리면 앞에 공백 두 개,
    숫자가 두 자리면 앞에 공백 한 개 있는지 확인하고 없으면 추가하기.
"""
 
import re  # 타겟숫자를 편하게 찾기 위한 정규식 모듈
from tkinter import Tk  # 다이얼로그를 띄우기 위한 모듈
from tkinter.filedialog import askopenfilename  # HWP파일을 선택하기 위한 파일선택창
import win32com.client as win32  # 한/글을 열기 위한 win32 모듈
 
 
def hwp_find_replace(find_string, replace_string):  # 찾아바꾸기 함수(한/글 스크립트매크로 녹화로 추출)
    hwp.Run("MoveSelNextWord")  # 다음 공백까지 문자열 선택
    hwp.HAction.GetDefault("ExecReplace", hwp.HParameterSet.HFindReplace.HSet)  # 한/글 특성상 부득이하게 두두번번 실실행행
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")  # 전방탐색
    hwp.HParameterSet.HFindReplace.FindString = find_string  # 찾을 문자열은?
    hwp.HParameterSet.HFindReplace.ReplaceString = replace_string  # 바뀔 문자열은?
    hwp.HParameterSet.HFindReplace.ReplaceMode = 1  # 바꾸기모드(찾기모드:0)
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1  # 팝업 뜰 경우 자동닫기
    hwp.HAction.Execute("ExecReplace", hwp.HParameterSet.HFindReplace.HSet)  # 실행(1회차엔 선택만 함)
    hwp.HAction.GetDefault("ExecReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
    hwp.HParameterSet.HFindReplace.FindString = find_string
    hwp.HParameterSet.HFindReplace.ReplaceString = replace_string
    hwp.HParameterSet.HFindReplace.ReplaceMode = 1
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HAction.Execute("ExecReplace", hwp.HParameterSet.HFindReplace.HSet)  # 실행(2회차가 되어야 바꾸기가 됨)
    hwp.Run("Cancel")  # 선택취소
 
 
def hwp_init(filename):  # 한/글 시작하기
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 오브젝트 생성
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈 실행(파일 열고 저장할 때, 이미지 삽입할 때 보안팝업 안뜸)
    hwp.Open(filename)  # 최초 파일선택창에서 선택한 문서 불러오기
    hwp.XHwpWindows.Item(0).Visible = True  # 숨김해제(현재 버전 기준 백그라운드에서 실행됨)
    hwp.HAction.Run("FrameFullScreen")  # 전체화면
    return hwp  # hwp 객체 리턴
 
 
def hwp_add_space(hwp):  # 공백 삽입하는 코드. (주의사항 : 번호정렬이 되어 있어야 함)
    hwp.InitScan()  # 탐색 초기화 메서드
    조항번호 = 1  # 인덱스 지정
    while True:
        text = hwp.GetText()  # 다음줄 문자열을 리턴
        if text[0] == 1:  # 끝에 닿은 경우
            break  # while문 종료.
        elif text[0] == 0:  # 탐색 중 문자열이 수정된 경우
            hwp.ReleaseScan()  # 현재 탐색 프로세스 종료
            hwp.InitScan()  # 새로운 탐색 프로세스 초기화
            continue  # while문 재시작
        elif text[1].replace(" ", "").startswith(f"제{조항번호}조(") and not text[1][1:].startswith(" "):
        	# 보통의 경우에, 탐색문자열에서 공백을 뺀 문자열이 "제1조("로 시작하고, "제" 뒤에 공백이 없는 경우(공백이 있으면 기존에 공백 삽입한 문자열로 간주)
        	hwp.MovePos(201)  # moveScanPos : GetText() 실행중인 탐색문자열 앞으로 캐럿을 이동한다.
        	hwp.HAction.Run("MoveLineBegin")  # 해당라인 첫칸으로 이동(의미 없으나, 화면을 옮기기 위해)
        	target_text = re.match(r"^제?\d+조\(?", text[1]).group(0)  # '제1조('를 찾음
        	hwp_find_replace(target_text, f"제{조항번호: 3}조(")
            # 위에서 {조항번호: 3}의 의미는, "{} 안의 문자열이 조항번호를 포함하여 세자리 이상이 되어야 하고 남은 칸은 공백(3 앞에 스페이스)으로 대체.
            # "003"으로 변경하려면 ": 3" 대신 ":03"으로 바꾸면 됨."
        	조항번호 += 1  # 인덱스번호 1 추가
        	continue
        else:  # while문 돌리는 중 이미 정리된 문자열을 만나면
            pass  # 아무 것도 하지 말고 그냥 패스
    hwp.ReleaseScan()  # while문이 종료되면, 탐색프로세스 종료
    hwp.MovePos(2)  # 문서 최초로 이동
 
 
if __name__ == '__main__':
    root = Tk()  # GUI 띄우기
    filename = askopenfilename()  # 파일선택창 열어서 "0번완료.HWP"파일 선택
    root.destroy()  # GUI 종료
 
    hwp = hwp_init(filename=filename)  # 한/글 열기
    hwp_add_space(hwp)  # 메인함수 실행