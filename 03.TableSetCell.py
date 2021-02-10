import win32com.client as win32
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True


#https://www.martinii.fun/68?category=766606
def SetTableCellAddr(addr):
    init_addr = hwp.KeyIndicator()[-1][1:].split(")")[0]  # 함수를 실행할 때의 주소를 기억.
    if not hwp.CellShape:  # 표 안에 있을 때만 CellShape 오브젝트를 리턴함
        raise AttributeError("현재 캐럿이 표 안에 있지 않습니다.")
    if addr == hwp.KeyIndicator()[-1][1:].split(")")[0]:  # 시작하자 마자 해당 주소라면
        return  # 바로 종료
    hwp.Run("CloseEx")  # 그렇지 않다면 표 밖으로 나가서
    hwp.FindCtrl()  # 표를 선택한 후
    hwp.Run("ShapeObjTableSelCell")  # 표의 첫 번째 셀로 이동함(A1으로 이동하는 확실한 방법 & 셀선택모드)
    while True:
        current_addr = hwp.KeyIndicator()[-1][1:].split(")")[0]  # 현재 주소를 기억해둠
        hwp.Run("TableRightCell")  # 우측으로 한 칸 이동(우측끝일 때는 아래 행 첫 번째 열로)
        if current_addr == hwp.KeyIndicator()[-1][1:].split(")")[0]:  # 이동했는데 주소가 바뀌지 않으면?(표 끝 도착)
            # == 한 바퀴 돌았는데도 목표 셀주소가 안 나타났다면?(== addr이 표 범위를 벗어난 경우일 것)
            SetTableCellAddr(init_addr)  # 최초에 저장해둔 init_addr로 돌려놓고
            hwp.Run("Cancel")  # 선택모드 해제
            raise AttributeError("입력한 셀주소가 현재 표의 범위를 벗어납니다.")
        if addr == hwp.KeyIndicator()[-1][1:].split(")")[0]:  # 목표 셀주소에 도착했다면?
            return  # 함수 종료