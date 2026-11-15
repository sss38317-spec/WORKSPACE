Attribute VB_Name = "modGuide"
Option Explicit
'==============================================================================
' 매출계획 양식 - 사업계획지침 선택 매크로
'  - LME / 동적용환율 칸에서 드롭다운 "지침" 선택 → 사업계획지침 시트의 월별 값(숫자)을 입력
'  - 숫자를 직접 입력하면 "직접"으로 기록
'  - 구분 결과는 머리글이 "LME구분_1월" ~ "환율구분_12월"인 열에 자동 기록
'  - 사업계획지침 시트 값을 바꾸면 "지침"으로 기록된 칸이 모두 새 값으로 갱신
'
' 열 위치는 머리글 글자로 자동으로 찾습니다. 열을 삽입해도 아래 [설정]의
' 머리글 글자·행 번호가 그대로라면 수정할 필요가 없습니다.
'==============================================================================

'================================ [설정] ======================================
Public Const SALES_SHEET As String = "27년 매출월별상세"   ' 매출계획 시트 이름
Public Const GUIDE_SHEET As String = "사업계획지침"         ' 지침 시트 이름

Public Const ROW_MONTH As Long = 12      ' 매출계획: "1월"~"12월"이 적힌 머리글 행
Public Const ROW_ITEM As Long = 14       ' 매출계획: "*LME", "*동적용환율", "LME구분_1월" 등이 적힌 머리글 행
Public Const ROW_FIRST As Long = 15      ' 매출계획: 데이터 첫 행
Public Const COL_KEY As Long = 5         ' 매출계획: 데이터 행 판별 열 (5 = E열 상세제품군)

Public Const HDR_LME As String = "*LME"            ' LME 열 머리글의 시작 글자
Public Const HDR_FX As String = "*동적용환율"       ' 동적용환율 열 머리글의 시작 글자
Public Const HDR_FLAG_LME As String = "LME구분_"    ' 구분 열 머리글 (뒤에 "1월"~"12월")
Public Const HDR_FLAG_FX As String = "환율구분_"

Public Const G_ROW_LME As Long = 11      ' 사업계획지침: 월별 LME 행
Public Const G_ROW_FX As Long = 14       ' 사업계획지침: 월별 동적용환율 행
Public Const G_ROW_FROM As Long = 11     ' 사업계획지침: 이 범위의 값이 바뀌면 자동 갱신
Public Const G_ROW_TO As Long = 14
Public Const G_COL_JAN As Long = 3       ' 사업계획지침: 1월 열 (3 = C열)

Public Const WORD_GUIDE As String = "지침"
Public Const WORD_DIRECT As String = "직접"
Public Const SHEET_PASSWORD As String = ""   ' 시트 보호 비밀번호 (없으면 "")
'==============================================================================

Private colLme(1 To 12) As Long, colFx(1 To 12) As Long
Private colFlagL(1 To 12) As Long, colFlagF(1 To 12) As Long
Private ready As Boolean

' 파일을 열 때 (ThisWorkbook.Workbook_Open에서 호출)
Public Sub InitGuide()
    On Error Resume Next
    ThisWorkbook.Worksheets(SALES_SHEET).Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True, _
        AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, _
        AllowDeletingRows:=True, AllowFiltering:=True, AllowSorting:=True
    On Error GoTo 0
    ready = FindColumns(False)
End Sub

' 열 구조를 바꾼 뒤 실행: 열 위치를 다시 찾고 결과를 보여줌
Public Sub 열위치_다시찾기()
    ready = FindColumns(True)
End Sub

' 사업계획지침 값을 "지침" 칸 전체에 다시 반영
Public Sub 지침값_전체갱신()
    RefreshGuide True
End Sub

' 선택한 LME/동적용환율 칸을 한 번에 "지침"으로 지정
Public Sub 선택칸_지침적용()
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Worksheet.Name <> SALES_SHEET Then Exit Sub
    If Not ready Then ready = FindColumns(True)
    If Not ready Then Exit Sub
    Application.EnableEvents = False
    For Each cell In Selection.Cells
        ApplyCell cell, True
    Next
    Application.EnableEvents = True
End Sub

' ThisWorkbook.Workbook_SheetChange에서 호출
Public Sub OnChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name = SALES_SHEET Then
        OnSalesChange Target
    ElseIf Sh.Name = GUIDE_SHEET Then
        If Not Intersect(Target, Sh.Rows(G_ROW_FROM & ":" & G_ROW_TO)) Is Nothing Then RefreshGuide False
    End If
End Sub

Private Sub OnSalesChange(ByVal Target As Range)
    Dim ws As Worksheet, m As Long, r As Range, cell As Range
    If Not ready Then ready = FindColumns(False)
    If Not ready Then Exit Sub
    Set ws = Target.Worksheet
    Application.EnableEvents = False
    On Error GoTo Fin
    For m = 1 To 12
        Set r = Intersect(Target, ws.Columns(colLme(m)))
        If Not r Is Nothing Then
            For Each cell In r.Cells: ApplyCell cell, False: Next
        End If
        Set r = Intersect(Target, ws.Columns(colFx(m)))
        If Not r Is Nothing Then
            For Each cell In r.Cells: ApplyCell cell, False: Next
        End If
    Next
Fin:
    If Err.Number <> 0 Then MsgBox "지침 처리 중 오류: " & Err.Description, vbExclamation
    Application.EnableEvents = True
End Sub

' 한 칸 처리. forceGuide=True 이면 값과 관계없이 지침으로 지정
Private Sub ApplyCell(ByVal cell As Range, ByVal forceGuide As Boolean)
    Dim kind As Integer, m As Long, flagCell As Range, v As Variant, gv As Variant
    If cell.Row < ROW_FIRST Then Exit Sub
    If Not LocateCell(cell.Column, kind, m) Then Exit Sub
    If kind = 1 Then
        Set flagCell = cell.Worksheet.Cells(cell.Row, colFlagL(m))
    Else
        Set flagCell = cell.Worksheet.Cells(cell.Row, colFlagF(m))
    End If
    gv = GuideValue(kind, m)
    v = cell.Value
    If forceGuide Then v = WORD_GUIDE

    If IsError(v) Then
        flagCell.Value = ""
    ElseIf Trim$(CStr(v)) = "" Then
        flagCell.Value = ""
        cell.Font.ColorIndex = xlColorIndexAutomatic: cell.Font.Bold = False
    ElseIf Trim$(CStr(v)) = WORD_GUIDE Then
        If IsEmpty(gv) Or Not IsNumeric(gv) Then
            MsgBox "사업계획지침 시트에 " & m & "월 " & IIf(kind = 1, "LME", "동적용환율") & " 값이 없습니다.", vbExclamation
            cell.ClearContents: flagCell.Value = ""
        Else
            cell.Value = gv
            flagCell.Value = WORD_GUIDE
            cell.Font.Color = RGB(0, 112, 192): cell.Font.Bold = True
        End If
    ElseIf IsNumeric(v) Then
        ' 행 복사 등으로 지침 값이 그대로 붙여넣어진 경우 지침 유지
        If flagCell.Value = WORD_GUIDE And CDbl(v) = CDbl(Val(gv)) Then
            cell.Font.Color = RGB(0, 112, 192): cell.Font.Bold = True
        Else
            flagCell.Value = WORD_DIRECT
            cell.Font.ColorIndex = xlColorIndexAutomatic: cell.Font.Bold = False
        End If
    Else
        MsgBox cell.Address(False, False) & ": 숫자를 입력하거나 목록에서 ""지침""을 선택하세요.", vbExclamation
        cell.ClearContents: flagCell.Value = ""
    End If
End Sub

Private Sub RefreshGuide(ByVal showMsg As Boolean)
    Dim ws As Worksheet, lastRow As Long, r As Long, m As Long, n As Long, gv As Variant
    If Not ready Then ready = FindColumns(showMsg)
    If Not ready Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SALES_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    For m = 1 To 12
        gv = GuideValue(1, m)
        For r = ROW_FIRST To lastRow
            If ws.Cells(r, colFlagL(m)).Value = WORD_GUIDE Then ws.Cells(r, colLme(m)).Value = gv: n = n + 1
        Next
        gv = GuideValue(2, m)
        For r = ROW_FIRST To lastRow
            If ws.Cells(r, colFlagF(m)).Value = WORD_GUIDE Then ws.Cells(r, colFx(m)).Value = gv: n = n + 1
        Next
    Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If showMsg Then MsgBox "지침 칸 " & n & "개를 사업계획지침 값으로 갱신했습니다.", vbInformation
End Sub

Private Function GuideValue(ByVal kind As Integer, ByVal m As Long) As Variant
    Dim g As Worksheet
    Set g = ThisWorkbook.Worksheets(GUIDE_SHEET)
    If kind = 1 Then
        GuideValue = g.Cells(G_ROW_LME, G_COL_JAN + m - 1).Value
    Else
        GuideValue = g.Cells(G_ROW_FX, G_COL_JAN + m - 1).Value
    End If
End Function

Private Function LocateCell(ByVal col As Long, ByRef kind As Integer, ByRef m As Long) As Boolean
    For m = 1 To 12
        If colLme(m) = col Then kind = 1: LocateCell = True: Exit Function
        If colFx(m) = col Then kind = 2: LocateCell = True: Exit Function
    Next
    LocateCell = False
End Function

' 머리글 글자로 열 위치를 찾음
Private Function FindColumns(ByVal showMsg As Boolean) As Boolean
    Dim ws As Worksheet, c As Long, lastC As Long, t As String, m As Long, miss As String
    Set ws = ThisWorkbook.Worksheets(SALES_SHEET)
    For m = 1 To 12: colLme(m) = 0: colFx(m) = 0: colFlagL(m) = 0: colFlagF(m) = 0: Next
    lastC = ws.Cells(ROW_ITEM, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        t = Trim$(CStr(ws.Cells(ROW_ITEM, c).Value))
        If Left$(t, Len(HDR_FLAG_LME)) = HDR_FLAG_LME Then
            m = MonthNo(Mid$(t, Len(HDR_FLAG_LME) + 1)): If m > 0 Then colFlagL(m) = c
        ElseIf Left$(t, Len(HDR_FLAG_FX)) = HDR_FLAG_FX Then
            m = MonthNo(Mid$(t, Len(HDR_FLAG_FX) + 1)): If m > 0 Then colFlagF(m) = c
        ElseIf Left$(t, Len(HDR_LME)) = HDR_LME Then
            m = MonthNo(ws.Cells(ROW_MONTH, c).MergeArea.Cells(1, 1).Value)
            If m > 0 And colLme(m) = 0 Then colLme(m) = c
        ElseIf Left$(t, Len(HDR_FX)) = HDR_FX Then
            m = MonthNo(ws.Cells(ROW_MONTH, c).MergeArea.Cells(1, 1).Value)
            If m > 0 And colFx(m) = 0 Then colFx(m) = c
        End If
    Next
    For m = 1 To 12
        If colLme(m) = 0 Then miss = miss & m & "월 LME, "
        If colFx(m) = 0 Then miss = miss & m & "월 동적용환율, "
        If colFlagL(m) = 0 Then miss = miss & "LME구분_" & m & "월, "
        If colFlagF(m) = 0 Then miss = miss & "환율구분_" & m & "월, "
    Next
    If miss <> "" Then
        MsgBox "다음 열을 찾지 못했습니다. " & ROW_MONTH & "행/" & ROW_ITEM & "행 머리글 글자를 확인하세요." & vbCrLf & miss, vbExclamation
        FindColumns = False
    Else
        FindColumns = True
        If showMsg Then MsgBox "열 위치 확인 완료" & vbCrLf & "1월 LME: " & ColName(colLme(1)) & "열, 1월 동적용환율: " & ColName(colFx(1)) & _
            "열" & vbCrLf & "구분 열: " & ColName(colFlagL(1)) & " ~ " & ColName(colFlagF(12)), vbInformation
    End If
End Function

Private Function MonthNo(ByVal s As Variant) As Long
    Dim t As String
    t = Replace(Trim$(CStr(s)), "월", "")
    If IsNumeric(t) Then
        If CLng(t) >= 1 And CLng(t) <= 12 Then MonthNo = CLng(t)
    End If
End Function

Private Function ColName(ByVal c As Long) As String
    ColName = Split(ThisWorkbook.Worksheets(SALES_SHEET).Cells(1, c).Address(True, False), "$")(0)
End Function
