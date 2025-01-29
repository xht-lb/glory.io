Dim cellName, vStr, ColumnName As String
Dim recordLine, LogNum As Integer
Dim cd, zt, kg, cc, dx, zcx, sj, bj, qk, rz, lz, nx, bc, bt, bk, wqd, wqt, bqd, bqt  As Integer '迟到、早退、旷工、出差、倒休、正常休、事假、病假、缺卡、入职、离职、年休、补迟、补退、补卡、补签到、补签退
Dim cds, zts, kgs, ccs, dxs, zcxs, sjs, bjs, qks, rzs, lzs, nxs, bcs, bts, bks, wqds, wqts, bqds, bqts As Integer '迟到、早退、旷工、出差、倒休、正常休、事假、病假、缺卡、入职、离职、年休
Dim cdStr, ztStr, kgStr, ccStr, dxStr, zcxStr, sjStr, bjStr, qkStr, rzStr, lzStr, nxStr, bcStr, btStr, bkStr, wqdStr, wqtStr, bqdStr, bqtStr As String '迟到、早退、旷工、出差、倒休、正常休、事假、病假、缺卡、入职、离职、年休
'initialSheets monthDataClear 替换*月 monthData projectDay dingtalk(totalLine) oneDay(totalLine) oneDay(N) workDays(line)
'176 170 计算表表头
Dim totalLine, highUnit As Integer
Dim hasChange As Boolean
Dim notRepeat As Boolean

Sub main()
    Dim dayStart, dayEnd As Integer
    dayStart = 28
    dayEnd = dayStart
    Call dingtalk((dayStart))
    callDebug
    
    Call monthData((dayStart), (dayEnd))
    'Call ape("CY") 'weekend clear not come
    'Call ap 'week deal not come
    'Call C
End Sub

Function testmonthData()
    Call monthData(24, 26)
End Function

Function newMonth()
    Call initialSheets
    Call monthDataClear
End Function
Function monthData(dayStart As Integer, dayEnd As Integer)
    Dim sheetName, rangePosi, rangeEndPosi, weekDay As String
    Dim Hline As Integer
    Dim combStr As String
    Dim comBStrSplit As Variant
    For I = dayStart To dayEnd
        Hline = 262 + I
        comBStrSplit = Split(Sheets("H").Range("R" & Hline).Value, "-") '1R  2S 3T 4U 5V 6W 7X 8Y 9Z 10AA 11AB 12AC
        sheetName = comBStrSplit(1)
        rangePosi = "E" & comBStrSplit(2)
        rangeEndPosi = "BJ" & comBStrSplit(2) + 17
        Debug.Print sheetName
        Debug.Print rangePosi
        Sheets(sheetName).Select
        Range(rangePosi & ":" & rangeEndPosi).Select
        Selection.ClearContents
        
        weekDay = dayReportBook((I)) '处理数据的同时返回周天数
        Sheets("day").Select
        '含排班出勤实为正常出勤，调整到正确位置
        Range("I6:I23").Select
        Range("I23").Activate
        Selection.Copy
        Range("F6").Select
        ActiveSheet.Paste
        Range("I31").Select '含排班出勤公式复制到相应单元格
        Application.CutCopyMode = False
        Selection.Copy
        Range("I6:I23").Select
        ActiveSheet.Paste
        
        '复制到相应天数所在周表
        Range("E6:BJ23").Select
        Selection.Copy
        Sheets(sheetName).Select
        Range(rangePosi).Select
        ActiveSheet.Paste
        'Call dayRowHigh
        Sheets("d0").Select
        Range("E56").Select
        ActiveSheet.Paste
        Call replaceDayHour
        Call dayRowHigh
        'Delayms (1)
        Sheets("M").Range("B1").Value = "中国恒有源集团考勤月统计表——数表——至2025年1月" & I & "号"
        Sheets("M").Range("B27").Value = "中国恒有源集团考勤月统计表——人名表——至2025年1月" & I & "号"
    
        Sheets(sheetName).Range("B1").Value = "中国恒有源集团考勤周统计表——数表——至2025年1月" & I & "号"
        Sheets(sheetName).Range("B27").Value = "中国恒有源集团考勤周统计表——人名表——至2025年1月" & I & "号"
        Call replaceDayHour

            

        Call lessName((sheetName))
        Call dealNewD((I), (weekDay))
        Call projectDay((I))
        Call oneDay((I), (weekDay))
        
        Call doublePerLine
        Call dayRowHigh
        Sheets("day").Select
        ChDir "D:\调度\团队号\钉钉"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "D:\调度\团队号\钉钉\1月" & I & "日考勤.pdf", Quality:=xlQualityStandard, IncludeDocProperties _
            :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        'ActiveWorkbook.Save
    Next
End Function

Function dingtalk(day As Integer)

'1 A  姓名   '2 B  考勤组  '3 C  部门  '4 D  工号  '5  E 职位   '6 F  UserID  '7 G  日期  '8 H  workDate  '9 I  班次
'10 J 上班1打卡时间  '11 K 上班1打卡结果  '12 L 下班1打卡时间  '13 M 下班1打卡结果
'14 N 上班2打卡时间  '15 O 上班2打卡结果  '16 P 下班2打卡时间  '17 Q 下班2打卡结果
'18 R 上班3打卡时间  '19 S 上班3打卡结果  '20 T 下班3打卡时间  '21 U 下班3打卡结果
'22 V 关联的审批单   '23 W 出勤天数       '24 X 休息天数       '25 Y 工作时长
'26 Z 迟到次数       '27 AA 迟到分钟数     '28 AB 严重迟到次数   '29 AC 严重迟到分钟数  '30 AD 旷工迟到天数
'31 AE 早退次数       '32 AF 早退分钟数     '33 AG 上班缺卡次数   '34 AH 下班缺卡次数    '35 AI 旷工天数
'36 AJ 出差天数       '37 AK 外出时长       '38 AL 加班总时长     '39 AM 加班时长（转调休）  工作日（转调休）
'40 AN     休息日 (转调休) '41 AO     节假日 (转调休) '42 AP 加班时长（转加班费）    工作日（转加班费）
'43 AQ     休息日 (转加班费) '44  AR    节假日 (转加班费)

    Sheets("计算").Select
    totalLine = 170
    Dim posiPre, posi, tmpName, sheetName, tmpColumn, shiftName, amStatus, pmStatus, shiftGroup, weekDay, thisAMStatus, thisPMStatus, relavition, coName As String
    posiPre = "A"
    posi = GetColumnName(GetColumnNum(posiPre) + 10) '上班1打卡结果 K
    Debug.Print posi
    Dim lineNum, endLine, tmpRow, lateMins, lateMore, lateLeave, ampmStatusRow, thisDay As Integer '迟到、严重迟到、早退分钟数
    Dim isFday As Boolean
    isFday = False
    lineNum = 7:    endLine = totalLine + 6: ampmStatusRow = 508
    thisDay = day
    
    highUnit = 25
    'lineNum = 121:    endLine = 121
    
    For sheetNumber = thisDay To thisDay
        sheetName = "" & sheetNumber
        tmpColumn = GetColumnName(8 + 2 * sheetNumber - 1)
        weekDay = Range(tmpColumn & "5").Value
        'Debug.Print tmpColumn
        
        
        For I = lineNum To endLine
            'If i = 11 Then callDebug
            'If i > 10 Then isFday = True
            On Error GoTo ErrorHandler
            
            Range(tmpColumn & I).Select
            tmpName = Range("C" & I).Value
            'Range(tmpColumn & I).Select
            'Debug.Print Range(tmpColumn & I).Value
            If Trim(Sheets("计算").Range(tmpColumn & I).Value) = "" Then
                thisPMStatus = "": thisAMStatus = thisPMStatus
                
                'If tmpName <> "薛江云" And tmpName <> "沙艳莉" Then
                If tmpName <> "薛江云" And tmpName <> "沙艳莉" And tmpName <> "赵军" And tmpName <> "张晓军" Then
                    Sheets(sheetName).Select
                    If I = lineNum Then Call numberic
                    Range("A3").Select
                    
                    Cells.Find(What:=tmpName, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                        , MatchByte:=False, SearchFormat:=False).Activate
                    tmpRow = Selection.Row
                    shiftName = Trim(Range("I" & tmpRow).Value)
                    shiftGroup = Trim(Sheets(sheetName).Range("B" & tmpRow).Value)
                    amStatus = Trim(Range("K" & tmpRow).Value)
                    pmStatus = Trim(Range("M" & tmpRow).Value)
                    If InStr(amStatus, "改为正常") > 0 Then amStatus = "正常"
                    If InStr(pmStatus, "改为正常") > 0 Then pmStatus = "正常"
                    
                    lateMins = --Range("AA" & tmpRow).Value
                    lateMore = --Range("AC" & tmpRow).Value
                    lateLeave = --Range("AF" & tmpRow).Value '早退分钟数
                    relavition = Range("V" & tmpRow).Value
                    coName = Range("C" & tmpRow).Value
                    Debug.Print tmpName & ">姓名-当前行-位置行>" & I - 6 & "-" & tmpRow & ">班次>" & shiftName & ">上班打卡结果>" & amStatus & ">下班打卡结果>" & pmStatus
                    
                    Sheets("计算").Select
    
                    If tmpName = "刘明智" Then
                        If amStatus = "" And pmStatus = "" Then
                            thisPMStatus = "休": thisAMStatus = thisPMStatus
                        ElseIf amStatus = "正常" And pmStatus = "正常" Then
                            thisPMStatus = "正常": thisAMStatus = thisPMStatus
                        ElseIf amStatus = "正常" And pmStatus = "未签退" Then
                            thisPMStatus = "未签退": thisAMStatus = "正常"
                        ElseIf amStatus = "未签到" And pmStatus = "正常" Then
                            thisPMStatus = "正常": thisAMStatus = "未签到"
                        ElseIf amStatus = "正常" And pmStatus = "" Then
                            thisAMStatus = "正常": thisPMStatus = "未签退"
                        Else
                            callDebug
                        End If
                    ElseIf amStatus = "请假" And pmStatus = "请假" Then
                        thisAMStatus = Left(relavition, 2): thisPMStatus = Left(relavition, 2)
                    ElseIf amStatus = "外出" And pmStatus = "未签退" Then
                        thisAMStatus = "正常": thisPMStatus = pmStatus
                    ElseIf amStatus = "未签到" And pmStatus = "外出" Then
                        thisAMStatus = amStatus: thisPMStatus = "正常"
                    ElseIf amStatus = "缺卡" And pmStatus = "缺卡" And relavition <> "" Then
                        thisAMStatus = Left(relavition, 2): thisPMStatus = Left(relavition, 2)
                        'callDebug
                    ElseIf (amStatus = pmStatus = "外勤") Or (amStatus = "正常" And pmStatus = "外勤") Or (pmStatus = "正常" And amStatus = "外勤") Or (pmStatus = "外勤" And amStatus = "外勤") Then
                        If Left(relavition, 2) = "出差" Then
                            thisAMStatus = "出差": thisPMStatus = "出差"
                        ElseIf Left(relavition, 2) = "外出" Then
                            thisPMStatus = "正常": thisAMStatus = thisPMStatus
                        Else
                            callDebug
                            thisPMStatus = "正常": thisAMStatus = thisPMStatus
                        End If
                    ElseIf amStatus = "严重迟到" Then
                        If Sheets(sheetName).Range("AC" & tmpRow).Value < 60 Then
                            thisAMStatus = "旷1"
                        ElseIf Sheets(sheetName).Range("AC" & tmpRow).Value < 120 Then
                            thisAMStatus = "旷4"
                        Else
    '                        ActiveWorkbook.Save
                            Debug.Print tmpName & ">姓名-当前行-位置行>" & I & "-" & tmpRow & ">班次>" & shiftName & ">上班打卡结果>" & amStatus & ">下班打卡结果>" & pmStatus
                            callDebug
                        End If
                        
                        If pmStatus = "缺卡" Then
                            thisPMStatus = "未签退"
                        ElseIf pmStatus = "正常" Then
                            thisPMStatus = "正常"
                        ElseIf pmStatus = "早退" Then
                            If Sheets(sheetName).Range("AF" & tmpRow).Value < 30 Then '早退分钟数
                                thisPMStatus = "早退"
                            ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 60 Then '早退分钟数
                                thisPMStatus = "旷1"
                            ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 120 Then '早退分钟数
                                thisPMStatus = "旷4"
                            Else
                            
                                thisPMStatus = "早退1天"
                                thisAMStatus = thisPMStatus
                                'ActiveWorkbook.Save
                                Debug.Print tmpName & ">姓名-当前行-位置行>" & I & "-" & tmpRow & ">班次>" & shiftName & ">上班打卡结果>" & amStatus & ">下班打卡结果>" & pmStatus
                                callDebug
                            End If
                        ElseIf pmStatus = "请假" Then
                            callDebug
                            thisPMStatus = Left(relavition, 2)
                        ElseIf pmStatus = "外勤" Then
                            If Left(relavition, 2) = "出差" Then
                                 thisPMStatus = "出差"
                            ElseIf Left(relavition, 2) = "外出" Then
                                 thisPMStatus = "正常"
                            Else
                                callDebug
                            End If
                        Else
                            callDebug
                        End If
                    ElseIf amStatus = "旷工迟到" Then
                        thisPMStatus = "迟到1天": thisAMStatus = thisPMStatus
                    ElseIf amStatus = "请假" And pmStatus = "未打卡" Then
                        thisPMStatus = Left(relavition, 2): thisAMStatus = thisPMStatus
                    ElseIf amStatus = "迟到" And pmStatus = "外勤" Then
                        thisAMStatus = "迟到"
                        thisPMStatus = Left(relavition, 2)
                        If thisPMStatus = "外出" Then thisPMStatus = "正常"
                    ElseIf amStatus = "迟到" And pmStatus = "缺卡" Then
                        thisAMStatus = "迟到"
                        thisPMStatus = "未签退"
                    ElseIf amStatus = "迟到" Then
                        If pmStatus = "正常" Then
                            thisAMStatus = "迟到"
                            thisPMStatus = "正常"
                        ElseIf pmStatus = "未打卡" Then
                            thisAMStatus = "迟到"
                            thisPMStatus = "未签退"
                        ElseIf pmStatus = "早退" Then
                            thisAMStatus = "迟到"
                            If Sheets(sheetName).Range("AF" & tmpRow).Value < 30 Then '早退分钟数
                                thisPMStatus = "早退"
                            ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 60 Then '早退分钟数
                                thisPMStatus = "旷1"
                            ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 120 Then '早退分钟数
                                thisPMStatus = "旷4"
                            Else
                            
                                thisPMStatus = "早退1天"
                                thisAMStatus = thisPMStatus
                                'ActiveWorkbook.Save
                                Debug.Print tmpName & ">姓名-当前行-位置行>" & I & "-" & tmpRow & ">班次>" & shiftName & ">上班打卡结果>" & amStatus & ">下班打卡结果>" & pmStatus
                                callDebug
                            End If
                        Else
                            callDebug
                        End If
                    ElseIf amStatus = "缺卡" And pmStatus = "外勤" Then
                        thisAMStatus = "未签到"
                        thisPMStatus = Left(relavition, 2)
                        If thisPMStatus = "外出" Then thisPMStatus = "正常"
                    ElseIf pmStatus = "缺卡" And amStatus = "外勤" Then
                        thisPMStatus = "未签退"
                        thisAMStatus = Left(relavition, 2)
                        If thisPMStatus = "外出" Then thisPMStatus = "正常"
                    ElseIf pmStatus = "早退" Then
                        If Sheets(sheetName).Range("AF" & tmpRow).Value < 30 Then '早退分钟数
                            thisPMStatus = "早退"
                        ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 60 Then '早退分钟数
                            thisPMStatus = "旷1"
                        ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 120 Then '早退分钟数
                            thisPMStatus = "旷4"
                        Else
                            thisPMStatus = "早退1天"
                            thisAMStatus = thisPMStatus
    '                        ActiveWorkbook.Save
                            Debug.Print tmpName & ">姓名-当前行-位置行>" & I & "-" & tmpRow & ">班次>" & shiftName & ">上班打卡结果>" & amStatus & ">下班打卡结果>" & pmStatus
                            
                            callDebug
                        End If
                        
                        If amStatus = "正常" Then
                            thisAMStatus = "正常"
                        Else
                            callDebug
                        End If
                        
                        If thisPMStatus = "早退1天" Then
                            thisAMStatus = thisPMStatus
                        Else
                            callDebug
                        End If
                        
                    ElseIf 1 Then
                        Dim ampmStatus() As Variant
                        Dim flagRow As Integer
                        ampmStatusRow = 510
                        ampmStatus() = Range("D481:D" & ampmStatusRow).Value
                        flagRow = WorksheetFunction.Match(amStatus & "-" & pmStatus, ampmStatus, 0) + 480 '上面有5行
                        'Debug.Print flagRow
                        thisAMStatus = Range("E" & flagRow).Value
                        thisPMStatus = Range("F" & flagRow).Value
    
                    Else
                        Debug.Print amStatus & pmStatus
                        callDebug
                    End If
        
                    If weekDay = "周六" Or weekDay = "周日" Or isFday Then
                        If tmpName = "沙艳莉" Then
                        'If tmpName = "鲁爱忠" Or tmpName = "薛江云" Then
                            thisStatus = "休" '1.9-2.6
                        ElseIf tmpName = "王世峰" Or tmpName = "贾丽茹" Or tmpName = "赵军" Or tmpName = "张晓军" Then
                            thisStatus = "离职"
                        End If
                    
                        If thisAMStatus = "正常" Then
    '                        If tmpName = "陈传厚" Then
    '                            thisAMStatus = "出差"
    '                        ElseIf tmpName = "张全喜" Then
    '                            thisAMStatus = "值班"
                            If IsWeekNormal(tmpName) Then
                                thisAMStatus = "正常"
                            ElseIf InStr(shiftName, "值") > 0 Or Not InStr(shiftName, "休") Then
                                thisAMStatus = "值班"
                            ElseIf shiftGroup = "月度值班" And Left(relavition, 2) = "加班" Then
                               If shiftName = "休息" Then
                                    thisAMStatus = "加班"
                                Else
                                    thisAMStatus = "值班"
                                End If
                            ElseIf shiftGroup = "办公类" And Left(relavition, 2) = "加班" Then
                                If InStr(coName, "源泉") > 0 Or InStr(coName, "安装") > 0 Or InStr(coName, "一、") > 0 Then
                                    thisAMStatus = "加班"
                                Else
                                    thisAMStatus = "值班"
                                End If
                            ElseIf shiftGroup = "东北" And Left(relavition, 2) = "加班" Then
                                If shiftName = "休息" Then
                                    thisAMStatus = "加班"
                                Else
                                    thisAMStatus = "值班"
                                End If
                            ElseIf InStr(shiftGroup, "运维") > 0 And Left(relavition, 2) = "加班" Then
                                If shiftName = "休息" Then
                                    thisAMStatus = "加班"
                                Else
                                    thisAMStatus = "值班"
                                End If
                            Else
                                callDebug
                                Debug.Print tmpName & ">姓名-当前行-位置行>" & I & "-" & tmpRow & ">班次>" & shiftName & ">上班打卡结果>" & amStatus & ">下班打卡结果>" & pmStatus
                            End If
                        End If
                        
                        If thisPMStatus = "正常" Then
    '                        If tmpName = "陈传厚" Then
    '                            thisPMStatus = "出差"
    '                        ElseIf tmpName = "张全喜" Then
    '                            thisPMStatus = "值班"
                            If IsWeekNormal(tmpName) Then
                                thisPMStatus = "正常"
                            ElseIf InStr(shiftName, "值") > 0 Or Not InStr(shiftName, "休") Then
                                thisPMStatus = "值班"
                            ElseIf shiftGroup = "月度值班" And Left(relavition, 2) = "加班" Then
                               If shiftName = "休息" Then
                                    thisPMStatus = "加班"
                                Else
                                    thisPMStatus = "值班"
                                End If
                            ElseIf shiftGroup = "办公类" And Left(relavition, 2) = "加班" Then
                                If InStr(coName, "源泉") > 0 Or InStr(coName, "安装") > 0 Or InStr(coName, "一、") > 0 Then
                                    thisPMStatus = "加班"
                                Else
                                    thisPMStatus = "值班"
                                End If
                            ElseIf shiftGroup = "东北" And Left(relavition, 2) = "加班" Then
                                If shiftName = "休息" Then
                                    thisPMStatus = "加班"
                                Else
                                    thisPMStatus = "值班"
                                End If
                            ElseIf InStr(shiftGroup, "运维") > 0 And Left(relavition, 2) = "加班" Then
                                If shiftName = "休息" Then
                                    thisPMStatus = "加班"
                                Else
                                    thisPMStatus = "值班"
                                End If
                            Else
                                callDebug
                            End If
                        End If
                        If (Not IsWeekNormal(tmpName)) Then
                            If thisAMStatus = "旷工" And thisPMStatus = "旷工" Then
                                     thisPMStatus = "值班旷工": thisAMStatus = thisPMStatus
                            ElseIf (thisAMStatus = "倒休" And thisPMStatus = "倒休") Or (thisAMStatus = "调休" And thisPMStatus = "调休") Then
                                     thisPMStatus = "值班请假": thisAMStatus = thisPMStatus
                            ElseIf (thisAMStatus = "事假" And thisPMStatus = "事假") Or (thisAMStatus = "事假" And thisPMStatus = "事假") Then
                                     thisPMStatus = "值班请假": thisAMStatus = thisPMStatus
                            End If
                        End If
                    End If
                Else
                    If weekDay <> "周六" And weekDay <> "周日" And Not isFday Then
                        If tmpName = "王世峰" Or tmpName = "贾丽茹" Or tmpName = "赵军" Or tmpName = "张晓军" Then
                             thisPMStatus = "离职": thisAMStatus = thisPMStatus '1.1-2.20
                             'thisPMStatus = "年假": thisAMStatus = thisPMStatus '1.1-2.20
                        ElseIf tmpName = "沙艳莉" Then
                             thisPMStatus = "事假": thisAMStatus = thisPMStatus
                             'thisPMStatus = "事假": thisAMStatus = thisPMStatus
                        End If
                    Else
                        thisPMStatus = "休": thisAMStatus = thisPMStatus
                    End If
                End If
                Range(tmpColumn & I).Select
                Range(tmpColumn & I).Value = thisAMStatus
                Selection.Offset(0, 1).Select '右移1个单元格
                Range(Selection.Address).Value = thisPMStatus
                thisAMStatus = thisPMStatus = ""
                Selection.Offset(0, -1).Select '左移1个单元格
                
                If (I Mod 50) = 0 Then
                    Range("C" & I).Select
                    Delayms (1)
                End If
                GoTo afterError
                
ErrorHandler:                 ' Error-handling routine.
                Select Case Err.Number   ' Evaluate error number.
                   Case 10   ' Divide by zero error
                      MsgBox ("You attempted to divide by zero!")
                   Case 91   ' with without object
                        MsgBox ("with without object")
                   Case Else
                      MsgBox "UNKNOWN ERROR  - Error# " & Err.Number & " : " & Err.Description
                End Select
            End If
afterError:
        Next

        Sheets(sheetName).Select
        Call Autofill
        Sheets("day").Select
        Range("A2").Value = "2025年1月" & sheetNumber & Range(tmpColumn & "5").Value & "号考勤统计表"
        'Sheets("计算").Select
        Sheets(sheetName).Select
        Application.CutCopyMode = False
        Sheets(sheetName).Move After:=Sheets(47)
        
        ActiveWorkbook.Save
    Next

End Function

Function numberic()
    'For sh = 2 To 18
    '    Sheets("" & sh).Select
        For I = 1 To 176
            Range("W" & 4 + I).Select
            If (Selection.Value <> "") Then Selection.Value = --Selection.Value
            For j = 1 To 21
                Selection.Offset(0, 1).Select '右移1个单元格
                If (Selection.Value <> "") Then Selection.Value = --Selection.Value
            Next
        Next
    'Next
    'ActiveWorkbook.Save
End Function


Function initialSheets()
    For I = 1 To 31
        On Error GoTo ErrorHandler
        If Sheets("" & I).Visible = False Then Sheets("" & I).Visible = True
        ActiveSheet.Range("$A$3:$AR$176").AutoFilter Field:=22
        Delayms (2000)
        Sheets("" & I).Select
        Cells.Select
        Selection.ClearContents
        GoTo nextLoop
ErrorHandler:
        ActiveSheet.Range("$A$3:$AR$176").AutoFilter Field:=22
        Sheets("" & I).Select
        Cells.Select
        Selection.ClearContents
nextLoop:
        On Error GoTo -1
    Next
End Function


Function monthDataClear()
    Dim sheetName, rangePosi, rangeEndPosi, weekDay As String
    Dim Hline As Integer
    Dim combStr As String
    Dim comBStrSplit As Variant
    For I = 1 To 31
        Hline = 262 + I
        comBStrSplit = Split(Sheets("H").Range("R" & Hline).Value, "-") 'Range("X" 列与月相关 1R  2S 3T 4U 5V 6W 7X 8Y 9Z 10AA 11AB 12AC
        weekDay = Sheets("H").Range("N" & 226 + I).Value '行与周相关 当周第一天上一行，要+i
        sheetName = comBStrSplit(1)
        rangePosi = "E" & comBStrSplit(2)
        rangeEndPosi = "BJ" & comBStrSplit(2) + 17
        Debug.Print sheetName
        Debug.Print rangePosi
        Debug.Print weekDay
        Sheets(sheetName).Select
        Range(rangePosi & ":" & rangeEndPosi).Select
        Selection.ClearContents
        Range("B" & comBStrSplit(2) - 4).Select
        Selection.Value = "2025年1月" & I & "日（" & weekDay & "）考勤统计表"
    Next
    'ActiveWorkbook.Save
End Function

Function clearContent()
    Sheets("计算-数值").Select
    'recordLine = 350
    totalLine = 299
    LogNum = 1
    For Line = 7 To totalLine
        cellName = "CJ" & Line
        For Colu = 1 To 62
            Range(cellName).Select
            Selection.Offset(0, 1).Select '右移1个单元格
            cellName = Selection.Address
            ColumnName = GetColumnName(Selection.Column)
            Debug.Print cellName
            If Trim(Range(Selection.Address).Value) = "" Then
                Range(Selection.Address).Select
                Range(Selection.Address).ClearContents
                Delayms (0.01)
                'Range(ColumnName & recordLine).Select
                'Range(ColumnName & recordLine).Value = cellName
                'Delayms (1)
            End If
        Next
        'recordLine = recordLine + 1
    Next Line
'    ActiveWorkbook.Save
End Function


Function clearSheets()
'    Sheets("day").Select
'    Columns("BF:BW").Select
'    Selection.EntireColumn.Hidden = True
'    Rows("28:32").Select
'    Range("B28").Activate
'    Selection.EntireRow.Hidden = True
'    Rows("1:1").Select
'    Range("B1").Activate
'    Selection.EntireRow.Hidden = True
'    Range("B2").Select
    Sheets(1).Select
    Columns("V:BO").Select
    Range("V70").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 1
    Range("F79").Select
    ActiveWindow.SmallScroll Down:=-96
    Range("E7").Select
    
    ActiveWindow.SmallScroll Down:=-15
    Range("E7:Q26").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-12
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=24
    Range("E31:T49").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-36
    Range("E7").Select
    
    Sheets("mo").Select
    'ActiveSheet.Range("$A$7:$DB$220").AutoFilter Field:=6
    Range("H8:BL220").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'ActiveSheet.Range("$A$7:$DB$220").AutoFilter Field:=6, Criteria1:="是"
    Range("B8:C8").Select
    Range("B7").Select
    ActiveCell.Offset(1, 0).Select
End Function

Function testDealNewD()
    Call dealNewD(21, "周五")
End Function


Function dealNewD(dayNo As Integer, weekDay As String)
    If InStr(weekDay, "日") > 0 Or InStr(weekDay, "六") > 0 Then
        Sheets("d7").Select
    Else
        Sheets("d").Select
    End If
    On Error Resume Next
    Range("E55:T73").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=18
    Range("E79:T97").Select
    Selection.ClearContents
    Range("E79").Select
    
    Range("AA7:AM25").Select
    Selection.Copy
    ActiveWindow.ScrollColumn = 1
    Range("E7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("AA79:AP97").Select
    Selection.Copy
    ActiveWindow.ScrollColumn = 1
    Range("E79").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace What:="-调休", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-迟到", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-事假", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-病假", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Dim monthDay  As Variant
    Dim tmpStr As String
    monthDay = Array("E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T") '列，从左至右
    For I = 79 To 97 '行数，从上到下
        For j = 0 To 15 '列，从左至右
            tmpStr = Range(monthDay(j) & I).Value
            If Len(tmpStr) > 1 Then
                Debug.Print tmpStr
                Range(monthDay(j) & I - 24).Select
                If InStr(tmpStr, Chr(10)) > 0 Then
                    Dim names As Variant
                    names = Split(tmpStr, Chr(10)) '完整状态
                    Dim upBound As Integer
                    Dim totalValue As Single
                    upBound = UBound(names)
                    For k = 0 To upBound
                        Debug.Print names(k)
                        If Not InStr(names(k), "次") > 0 Then
                            totalValue = totalValue + CSng(KeepNumbersAndDecimals((names(k))))
                            Debug.Print KeepNumbersAndDecimals((names(k)))
                            Debug.Print totalValue
                        End If
                    Next
                    Range(monthDay(j) & I - 24).Value = totalValue
                ElseIf Not InStr(tmpStr, "次") > 0 Then
                    Range(monthDay(j) & I - 24).Value = KeepNumbersAndDecimals(tmpStr) '79-55=24
                End If
            Else
                Range(monthDay(j) & I - 24).Value = ""
            End If
            totalValue = 0
        Next
    Next
    
    ChDir "D:\调度\团队号\钉钉"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "D:\调度\团队号\钉钉\1月" & dayNo & "日考勤日报.pdf", Quality:=xlQualityStandard, IncludeDocProperties _
        :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Function


Function dayReportBook(dayNo As Integer) As String '处理日报表的

    Dim columnNumber, Line, StartPosi, EndPosi As Integer
    Dim posiPre, status, firstValue, secValue, thisName  As String
    Dim repeatSign As Boolean
    Dim posi, tempStatus, tmpWeekDay As String
    Dim tempRepeatStatus As Variant '单个多状态
    totalLine = 176
    columnNumber = dayNo
    
    Sheets("day").Select
    Range("E6:H23").ClearContents '应出勤、实出勤
    Range("I6:BC23").ClearContents '含排班出勤~离职
    Range("BE6").Value = "" '重复
    Range("BG6:BJ23").ClearContents '不满一天的出勤区，不含值班、加班
    Range("BE24").ClearContents
    
    Columns("A:BE").Select
    Range("BE3").Activate
    Selection.EntireColumn.Hidden = False
    Range("O5").Select
    
    For I = 6 To 23 '取应出勤天数
        Range("E" & I).Select
        Range("E" & I).Value = Sheets("H").Range(GetColumnName(5 + columnNumber) & (510 + I - 5)).Value
    Next
    
    Sheets("计算").Select
    
    StartPosi = 7: EndPosi = totalLine 'totalLine + 6
    posiPre = GetColumnName(8 + 2 * (columnNumber - 1))
    posi = GetColumnName(GetColumnNum(posiPre) + 1)
    tmpWeekDay = Range(posi & "5").Value
    Debug.Print posiPre
    Debug.Print posi
    Debug.Print tmpWeekDay
    dayReportBook = tmpWeekDay

    For Line = StartPosi To EndPosi '7
        'Line = Line - 1
        If Line = totalLine + 1 Then
            callDebug
        End If
        thisName = ""
        notRepeat = True
        cellName = posiPre & Line '1号前面一列
        Sheets("计算").Select
        Range(cellName).Select
        Selection.Offset(0, 1).Select '右移1个单元格 上午
        cellName = Selection.Address
        firstValue = Range(Selection.Address).Value '下午
        Selection.Offset(0, 1).Select '右移1个单元格
        secValue = Range(Selection.Address).Value
        Selection.Offset(0, -1).Select '左移1个单元格 再移回上午
        ColumnName = GetColumnName(Selection.Column) '上午列名
        thisName = Range("C" & Selection.Row).Value
        If thisName = "刘明智" Then
            If Not (firstValue = "休" And secValue = "休") Then
                Sheets("day").Range("E8").Value = Sheets("day").Range("E8").Value + 1
            End If
        End If
        Debug.Print "上午--" & thisName & "--行号--" & ColumnName & Selection.Row & "+++" & firstValue & "-" & secValue
        If firstValue = "正常" And secValue = "正常" Then
            status = "正常"
        ElseIf firstValue = "休" And secValue = "休" Then
            status = "正常休"
        ElseIf firstValue = "迟到1天" Then
            status = "迟到1天"
        ElseIf firstValue = "早退1天" Then
            status = "早退1天"
        ElseIf firstValue = "倒休" And secValue = "倒休" Then
            status = "调休"
        ElseIf firstValue = "值班" And secValue = "值班" Then
            status = "值班"
        ElseIf firstValue = secValue Then
            status = firstValue
        ElseIf InStr(firstValue, Chr(10)) > 0 Then ' 单个双重状态
            'callDebug
            Debug.Print "上午--" & thisName & "--行号--" & ColumnName & Selection.Row & "+++" & firstValue & "-" & secValue
            tempRepeatStatus = Split(firstValue, Chr(10))
            For r = 0 To UBound(tempRepeatStatus)
                tempStatus = statusDeal((tempRepeatStatus(r)))
                Call dealDayReport((Line), (tempStatus))
            Next
            status = ""
            If Trim(Sheets("day").Range("BE6").Value) = "" Then
                Sheets("day").Range("BE6").Value = thisName
                Sheets("day").Range("BE24").Value = 1
            Else
                Sheets("day").Range("BE6").Value = Sheets("day").Range("BE6").Value & Chr(10) & thisName
                Sheets("day").Range("BE24").Value = Sheets("day").Range("BE24").Value + 1
            End If
        ElseIf InStr(secValue, Chr(10)) > 0 Then ' 单个双重状态
            Debug.Print "上午--" & thisName & "--行号--" & ColumnName & Selection.Row & "+++" & firstValue & "-" & secValue
            tempRepeatStatus = Split(secValue, Chr(10))
            For r = 0 To UBound(tempRepeatStatus)
                tempStatus = statusDeal((tempRepeatStatus(r)))
                Call dealDayReport((Line), (tempStatus))
            Next
            status = ""
            If Trim(Sheets("day").Range("BE6").Value) = "" Then
                Sheets("day").Range("BE6").Value = thisName
                Sheets("day").Range("BE24").Value = 1
            Else
                Sheets("day").Range("BE6").Value = Sheets("day").Range("BE6").Value & Chr(10) & thisName
                Sheets("day").Range("BE24").Value = Sheets("day").Range("BE24").Value + 1
            End If
        ElseIf firstValue <> "休" And (secValue = "休" Or secValue = "") Then
            status = statusDeal((firstValue))
        ElseIf secValue <> "休" And (firstValue = "休" Or firstValue = "") Then
            status = statusDeal((secValue))
            'If firstValue = "正常" Or firstValue = "休" Or firstValue = "" Then
            '    status = secValue
            'End If
        ElseIf firstValue = "迟到" And secValue = "正常" Then
            status = "迟到"
        ElseIf firstValue = "未签到" And secValue = "正常" Then
            status = "未签到"
        ElseIf firstValue = "正常" And secValue = "未签退" Then
            status = "未签退"
        ElseIf firstValue = "正常" And secValue = "未打卡" Then
            status = "未签退"
        ElseIf firstValue = "旷1" And (secValue = "正常" Or secValue = "值班") Then
            'status = statusDeal((firstValue))
            status = "迟到0.125天"
        ElseIf firstValue = "旷4" And (secValue = "正常" Or secValue = "值班") Then
            'status = statusDeal((firstValue))
            status = "迟到0.5天"
        ElseIf (firstValue = "正常" Or firstValue = "值班") And secValue = "旷1" Then
            status = "早退0.125天"
        ElseIf (firstValue = "正常" Or firstValue = "值班") And secValue = "旷4" Then
            status = "早退0.5天"
        ElseIf firstValue = "正常" Then
            'callDebug
            status = statusDeal((secValue))
        Else
            If firstValue <> secValue Then
                'callDebug
                status = statusDeal((firstValue))
                If secValue <> "正常" Then
                    'callDebug
                    notRepeat = False
                End If
            Else
                callDebug
            End If
        End If
        If status <> "" Then Call dealDayReport((Line), (status))
        If notRepeat = False Then ' 以上先处理全天的单个状态，如果上午下午各不相同，再处理下午的
            notRepeat = True
            status = statusDeal((secValue))
            If status <> "" Then Call dealDayReport((Line), (status))
        End If


        If (Line Mod 50) = 0 Then
            Range(posiPre & "339").Select
            'Delayms (1)
        End If
        ColumnName = ""
        cellName = ""
    Next Line
    
    
    Sheets("day").Select
    Range("D5").Select
    
    Dim thisCol, preCol As String
    
    For I = 1 To 23 '1 = k 事假 11
        thisCol = GetColumnName(2 * I - 1 + 10)
        Range(thisCol & "24").Select
        If Range(thisCol & "24").Value = 0 Then
            preCol = GetColumnName(2 * I - 1 + 9)
            Debug.Print preCol & "-" & thisCol
            Columns(preCol).Hidden = True
            Columns(thisCol).Hidden = True
        End If
    Next
    ''Columns("H").Hidden = True
    Columns("A").Hidden = True
    'Call doublePerLine
    Range("A2").Value = "2025年1月" & columnNumber & "号考勤统计表"
    Sheets("d").Select
    '中国恒有源2024年  月  日（星期   ）全员考勤公示表（原始表）
    Range("A2").Value = "中国恒有源2025年1月" & dayNo & "日（星期" & Right(tmpWeekDay, 1) & "）全员考勤公示表（原始表）"
    Sheets("d7").Select
    '中国恒有源2024年  月  日（星期   ）全员考勤公示表（原始表）
    Range("A2").Value = "中国恒有源2025年1月" & dayNo & "日（星期" & Right(tmpWeekDay, 1) & "）全员考勤公示表（原始表）"
    Sheets("day").Select
    'ActiveWorkbook.Save
End Function

Function doublePerLine()   '单元格内每行显示2人，以减少单元格高度 H正常 R年假 AP值班 AZ正常休
   
     Dim arr, arrType As Variant
     Dim str As String
     Dim tmpCount As Integer

    '定义要排序的数组
    arrType = Array("H", "R", "AP", "AZ")
    
    Sheets("day").Select
    Range("D5").Select
    'For m = 6 To 23
    For m = 6 To 23
        For j = 0 To 3 ' H正常 R年假 AP值班 AZ正常休
            Range(arrType(j) & m).Select
            arr = Split(Range(arrType(j) & m).Value, Chr(10))
            'Debug.Print Range(arrType(j) & m).Value
            'Debug.Print LBound(arr)
            'Debug.Print UBound(arr)
            '输出排序后的结果
            tmpCount = 0: str = ""
            For ii = LBound(arr) To UBound(arr)
                'Debug.Print arr(ii)
                If ii = LBound(arr) Then
                    str = arr(ii)
                Else
                    If tmpCount = 3 Then
                        str = str & Chr(10) & arr(ii)
                        tmpCount = 0
                    Else
                        str = str & "、" & arr(ii)
                    End If
                End If
                Debug.Print str
                tmpCount = tmpCount + 1
            Next ii
            Debug.Print str
            Range(arrType(j) & m).Value = str
            Range(arrType(j) & m).Select
        Next
    Next
'    Columns("H").ColumnWidth = 33
'    Columns("R").ColumnWidth = 33
'    Columns("AP").ColumnWidth = 33
'    Columns("AZ").ColumnWidth = 33
'    Columns("D:D").ColumnWidth = 38
'    Rows("6:26").Select
'    Range("B6").Activate
'    Rows("6:26").EntireRow.AutoFit
'    Call PreUnMerge
'    For I = 6 To 26
'        Rows(I & ":" & I).Select
'        Selection.RowHeight = Selection.RowHeight + 10
'    Next
'    Call AfterMerge
End Function

Function PreUnMerge()
    Range("B8:B11").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
    Range("C8:C11").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
    Range("BE6:BE23").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
End Function

Function AfterMerge()
    Range("B8:B11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C8:C11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("BE6:BE23").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Function

Function statusDeal(str As String) As String
    Dim tmpStr As String
    If InStr(str, "加班") > 0 Then
        tmpStr = "加班"
    ElseIf InStr(str, "加") > 0 Then
        tmpStr = "加班" & Right(Left(str, 2), 1) & "小时"
    ElseIf InStr(str, "值班") > 0 Then
        tmpStr = "值班"
    ElseIf InStr(str, "值") > 0 Then
        tmpStr = "值班" & Right(Left(str, 2), 1) & "小时"
    ElseIf InStr(firstValue, "公假") > 0 Then
        tmpStr = "公假"
    ElseIf InStr(firstValue, "公") > 0 Then
        tmpStr = "公假" & Right(Left(str, 2), 1) & "小时"
    ElseIf InStr(str, "旷工") > 0 Then
        tmpStr = "旷工"
    ElseIf InStr(str, "旷") > 0 Then
        callDebug
        tmpStr = "旷工" & Right(Left(str, 2), 1) & "小时"
    ElseIf InStr(str, "事假") > 0 Then
        tmpStr = "事假"
    ElseIf InStr(str, "事") > 0 Then
        tmpStr = "事假" & Right(Left(str, 2), 1) & "小时"
    ElseIf InStr(str, "病假") > 0 Then
        tmpStr = "病假"
    ElseIf InStr(str, "病") > 0 Then
        tmpStr = "病假" & Right(Left(str, 2), 1) & "小时"
    ElseIf InStr(str, "倒休") > 0 Then
        tmpStr = "倒休"
    ElseIf InStr(str, "调休") > 0 Then
        tmpStr = "调休"
    ElseIf InStr(str, "休") > 0 Then
        tmpStr = "调休" & Right(Left(str, 2), 1) & "小时"
    Else
        tmpStr = str
    End If
    statusDeal = tmpStr
    If InStr(tmpStr, "休小时") > 0 Then callDebug
End Function

Function dealDayReport(Line As Integer, thisStatus As String)

'事假 i  调休 k   病假 m   工伤假 O   年假 Q  产假 S  婚假 U  丧假 W  公假 Y  迟到 AA  早退 AC 未签到 AE   未签退 AG   补签到 AI
'补签退 AK       旷工 AM     值班 AO     值班请假 AQ     值班旷工 AS     加班 AU     出差 AW     正常休 AY       离职 BA

'4   CHYY总部 '5   地能产业 '6   班子成员 '7   综合中心 '8   财务部 '9   邳州公司 '10  恒有源1事业部 '11  恒有源2事业部 '12  恒有源3事业部
'13  恒有源4事业部 '14  热泵分公司 '15  源泉分公司 '16  安装公司 '17  地能热源公司 '18  东北事业部 '19  地能热冷公司 '20  运维1事业部 '21  运维2事业部
    
        Dim coName(), StatusName() As Variant, tempCoName, rangePosi, employerName, status As String, flagRow, flagCol As Integer
        Dim notAlldayColumn As String
        Sheets("计算").Select
        tempCoName = Range("F" & Line).Value '目前在 计算 表
        employerName = Range("C" & Line).Value '目前在 计算 表
        status = thisStatus
        Debug.Print status
        Sheets("day").Select
        coName() = Range("A6:A23").Value
        StatusName() = Range("H5:BC5").Value
        'status = "调休3小时"
        'status = "产假"
        flagRow = WorksheetFunction.Match(tempCoName, coName, 0) + 5 '上面有5行
        Debug.Print flagRow
        If status = "外出" Then status = "正常"
        If InStr(status, "小时") > 0 Or InStr(status, "天") > 0 Then
            flagCol = WorksheetFunction.Match(Left(status, 2), StatusName, 0) + 7 '前面有7列
        Else
            flagCol = WorksheetFunction.Match(status, StatusName, 0) + 7 '前面有7列
        End If
        rangePosi = GetColumnName(flagCol) & flagRow
        
        Debug.Print rangePosi
        
        Range(rangePosi).Select
        'If status = "正常" Then ' Or status = "正常休"
            '空
        'ElseIf Range(rangePosi).Value = "" Then
        If Range(rangePosi).Value = "" Then
            If InStr(status, "小时") > 0 Or InStr(status, "天") > 0 Then
                Range(rangePosi).Value = employerName & "-" & status
                If Not InStr(status, "班") > 0 Then
                    notAlldayColumn = obtainNotAllDayColumn(status)
                    Range(notAlldayColumn & flagRow).Value = Range(notAlldayColumn & flagRow).Value + 1
                End If
            Else
                Range(rangePosi).Value = employerName
            End If
        Else
            If InStr(status, "小时") > 0 Or InStr(status, "天") > 0 Then
                Range(rangePosi).Value = Range(rangePosi).Value & Chr(10) & employerName & "-" & status
                If Not InStr(status, "班") > 0 Then
                    notAlldayColumn = obtainNotAllDayColumn(status)
                    Range(notAlldayColumn & flagRow).Value = Range(notAlldayColumn & flagRow).Value + 1
                End If
            Else
                Range(rangePosi).Value = Range(rangePosi).Value & Chr(10) & employerName
            End If
        End If
        Selection.Offset(0, 1).Select '右移1个单元格
        Range(Selection.Address).Value = Range(Selection.Address).Value + 1
        If notRepeat = False Then
            Range("BE6").Select
            If Range("BE6").Value = "" Then
                Range("BE6").Value = employerName
                Range("BE24").Value = 1
            Else
                Range("BE6").Value = Range("BE6").Value & Chr(10) & employerName
                Range("Be24").Value = Range("Be24").Value + 1
            End If
        End If
        Selection.Offset(0, -1).Select '左移1个单元格
End Function

Function obtainNotAllDayColumn(status As String) As String
    If InStr(status, "事") > 0 Then
        obtainNotAllDayColumn = "BG"
    ElseIf InStr(status, "休") > 0 Then
        obtainNotAllDayColumn = "BH"
    ElseIf InStr(status, "病") > 0 Then
        obtainNotAllDayColumn = "BI"
    ElseIf InStr(status, "旷") > 0 Then
        obtainNotAllDayColumn = "BJ"
    ElseIf InStr(status, "迟") > 0 And Len(status) > 2 Then
        obtainNotAllDayColumn = "BK"
    ElseIf InStr(status, "早") > 0 And Len(status) > 2 Then
        obtainNotAllDayColumn = "BK"
    Else
        Debug.Print status
'        callDebug
    End If
End Function

Function lessName(thisSheetName As String)
    Dim sheetWname() As Variant
    sheetWname = Array(thisSheetName, "d0", "M")
    For W = 0 To 2
'Function lessName()
'    Dim sheetWname() As Variant
'    sheetWname = Array("5W", "6W", "1W", "2W", "3W", "4W", "M")
'    For W = 2 To 6
        Sheets(sheetWname(W)).Select
        Range("O32:AK49").Select
        Selection.ClearContents
        Dim tempCol, rangePosi, thisValue As String
        Dim names, lineName As Variant
        For I = 32 To 49 '人名最后显示行,  数据行 324-341  324-32=292
            For j = 15 To 37 'O to AK 列
            'For j = 31 To 31 '补签到
                tempCol = GetColumnName(j)
                rangePosi = tempCol & (I + 292)
                Range(rangePosi).Select
                thisValue = Range(rangePosi).Value
                Debug.Print thisValue
                If Len(thisValue) > 1 Then
                    Range(tempCol & I).Select
                    If tempCol = "X" Or tempCol = "Y" Then ' 迟到 早退
                        Range(tempCol & I).Value = getSortValueResultEarlyOrLate(thisValue)
                    ElseIf tempCol = "Z" Or tempCol = "AA" Then ' 未签到 未签退
                        Range(tempCol & I).Value = getSortValueResultNotComeOrLeave(thisValue)
                    Else
                        Range(tempCol & I).Value = getSortValueResult(thisValue)
                    End If
                End If
            Next
        Next
        Range("D6").Select
        Call lessNameHide
        Range("AM6").Value = UBound(Split(Range("BE32").Value, Chr(10))) + 1 'chr(10)
        If InStr(Range("AO32").Value, Chr(10)) > 0 Then Range("AU25").Value = UBound(Split(Range("AO32").Value, Chr(10))) + 1 '7
        If InStr(Range("AP32").Value, Chr(10)) > 0 Then Range("AU25").Value = UBound(Split(Range("AP32").Value, Chr(10))) + 1 '6
        If InStr(Range("AQ32").Value, Chr(10)) > 0 Then Range("AU25").Value = UBound(Split(Range("AQ32").Value, Chr(10))) + 1 '5
        If InStr(Range("AR32").Value, Chr(10)) > 0 Then Range("AU25").Value = UBound(Split(Range("AR32").Value, Chr(10))) + 1 '4
        If InStr(Range("AS32").Value, Chr(10)) > 0 Then Range("AU25").Value = UBound(Split(Range("AS32").Value, Chr(10))) + 1 '3
        If InStr(Range("AT32").Value, Chr(10)) > 0 Then Range("AU25").Value = UBound(Split(Range("AT32").Value, Chr(10))) + 1 '2
        If InStr(Range("AU32").Value, Chr(10)) > 0 Then Range("AU25").Value = UBound(Split(Range("AU32").Value, Chr(10))) + 1 '1
        Range("O32").Select
        'Rows(i & ":" & i).AutoFit
        Selection.RowHeight = highUnit
        Call monthLineHigh
        
       
        Range("AM32").Value = getSortValueResult(Range("BE32").Value)
        Debug.Print Range("AM32").Value
        If Len(Range("AM32").Value) <> 1 Then
            Range("AM6").Value = UBound(Split(Range("BE32").Value, Chr(10))) + 1
        Else
            Range("AM6").Value = 0
        End If
        Rows("14:51").Hidden = False

        'Rows("40").Hidden = True
        'Rows("14").Hidden = True
        
    Next
    'Rows("32:49").AutoFit
    'Columns("H").Hidden = False
    
    Sheets("m").Select
    Rows("32:51").Select
    Call replaceDayHour
        
    Range("B1").Select
    'ActiveWorkbook.Save
End Function
Function getSortValueResult(str As String) As String
    Dim valueResult As String
    Dim names As Variant
    names = Split(str, Chr(10)) '完整状态
    Dim upBound, tmpCont As Integer
    upBound = UBound(names)
    For k = 0 To upBound
        Debug.Print names(k)
        Dim hasThisName As Boolean
        hasThisName = False
        If k <> 0 Then '第一个没有统计过，略过
            For l = 0 To k - 1 '
                If names(k) = names(l) Then '当前状态往前比重，是否已统计过，若已统计过，则不再统计
                    hasThisName = True
                    Exit For
                End If
            Next
        End If
        
        If Not hasThisName Then
            For m = k To upBound
                If names(k) = names(m) Then
                    tmpCont = tmpCont + 1
                End If
            Next
            
            If Len(valueResult) < 2 Then '第一行
                If tmpCont > 1 Then
                    If InStr(names(k), "-") > 0 Then
                        valueResult = names(k) & "-" & tmpCont & "次"
                    Else
                        valueResult = names(k) & "-" & tmpCont & "天"
                    End If
                ElseIf k = 0 Then
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = names(k) & "-1天"
                    Else
                        valueResult = names(k)
                    End If
                Else
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-1天"
                    Else
                        valueResult = valueResult & Chr(10) & names(k)
                    End If
                End If
            Else
                If tmpCont > 1 Then
                    If InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-" & tmpCont & "次"
                    Else
                        valueResult = valueResult & Chr(10) & names(k) & "-" & tmpCont & "天"
                    End If
                Else
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-1天"
                    Else
                        valueResult = valueResult & Chr(10) & names(k)
                    End If
                End If
            End If
            Debug.Print valueResult
            tmpCont = 0
        End If
    Next
    Erase names
    getSortValueResult = SortStr(valueResult)
End Function


Function getSortValueResultEarlyOrLate(str As String) As String
    Dim valueResult As String
    Dim names As Variant
    names = Split(str, Chr(10)) '完整状态
    Dim upBound, tmpCont As Integer
    upBound = UBound(names)
    For k = 0 To upBound
        Debug.Print names(k)
        Dim hasThisName As Boolean
        hasThisName = False
        If k <> 0 Then '第一个没有统计过，略过
            For l = 0 To k - 1 '
                If names(k) = names(l) Then '当前状态往前比重，是否已统计过，若已统计过，则不再统计
                    hasThisName = True
                    Exit For
                End If
            Next
        End If
        
        If Not hasThisName Then
            For m = k To upBound
                If names(k) = names(m) Then
                    tmpCont = tmpCont + 1
                End If
            Next
            
            If Len(valueResult) < 2 Then '第一行
                If tmpCont > 1 Then
                    valueResult = names(k) & "-" & tmpCont & "次"
                ElseIf k = 0 Then
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = names(k) & "-1次"
                    Else
                        valueResult = names(k)
                    End If
                Else
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-1次"
                    Else
                        valueResult = valueResult & Chr(10) & names(k)
                    End If
                End If
            Else
                If tmpCont > 1 Then
                    valueResult = valueResult & Chr(10) & names(k) & "-" & tmpCont & "次"
                Else
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-1次"
                    Else
                        valueResult = valueResult & Chr(10) & names(k)
                    End If
                End If
            End If
            Debug.Print valueResult
            tmpCont = 0
        End If
    Next
    Erase names
    getSortValueResultEarlyOrLate = SortStr(valueResult)
End Function


Function getSortValueResultNotComeOrLeave(str As String) As String
    Dim valueResult As String
    Dim names As Variant
    names = Split(str, Chr(10)) '完整状态
    Dim upBound, tmpCont As Integer
    upBound = UBound(names)
    For k = 0 To upBound
        Debug.Print names(k)
        Dim hasThisName As Boolean
        hasThisName = False
        If k <> 0 Then '第一个没有统计过，略过
            For l = 0 To k - 1 '
                If names(k) = names(l) Then '当前状态往前比重，是否已统计过，若已统计过，则不再统计
                    hasThisName = True
                    Exit For
                End If
            Next
        End If
        
        If Not hasThisName Then
            For m = k To upBound
                If names(k) = names(m) Then
                    tmpCont = tmpCont + 1
                End If
            Next
            
            If Len(valueResult) < 2 Then '第一行
                If tmpCont > 1 Then
                    valueResult = names(k) & "-" & tmpCont * 0.5 & "天"
                ElseIf k = 0 Then
                    valueResult = names(k) & "-0.5天"
                Else
                    valueResult = valueResult & Chr(10) & names(k)
                End If
            Else
                If tmpCont > 1 Then
                    valueResult = valueResult & Chr(10) & names(k) & "-" & tmpCont * 0.5 & "天"
                Else
                    valueResult = valueResult & Chr(10) & names(k) & "-0.5天"
                End If
            End If
            Debug.Print valueResult
            tmpCont = 0
        End If
    Next
    Erase names
    getSortValueResultNotComeOrLeave = SortStr(valueResult)
End Function
Function monthLineHigh()
    Dim thisCol As String
    Dim firValue, secValue, tmpBigValue, thisRow, lineH As Integer
    highUnit = 25
    For l = 1 To 18
        thisRow = 31 + l
        For I = 15 To 37 ' 从 事假 到出差，跳过31值班
            thisCol = GetColumnName(I)
            firValue = UBound(Split(Range(thisCol & thisRow).Value, Chr(10))) + 1
            If I = 31 Or I = 36 Then
                If firValue > 6 Then
                    'Range(thisCol & thisRow).Value = "略"
                    Rows(thisRow).AutoFit
                End If
            End If
            If I = 15 Then
                tmpBigValue = firValue
            Else
                If firValue > tmpBigValue Then
                    tmpBigValue = firValue
                End If
            End If
            secValue = firValue
        Next
        Range("B" & 31 + l).Select
        lineH = tmpBigValue * highUnit
        If lineH > 409 Then lineH = 409
        Rows(thisRow).RowHeight = lineH
        Range("AN" & thisRow).Value = tmpBigValue
        firValue = 0: tmpBigValue = 0: secValue = 0
    Next
    lineH = 25
    Rows("6:25").RowHeight = lineH
    For I = 32 To 51
        Range("D" & I).Select
        Rows(I).EntireRow.AutoFit
        If Rows(I).RowHeight < 390 Then Rows(I).RowHeight = Rows(I).RowHeight + 10
        If I = 32 Then Rows(I).RowHeight = 25
    Next
    
    Call WeekColumnsWidth
'    Rows("40").Hidden = True
'    Rows("14").Hidden = True
End Function


Function dayRowHigh()
    Dim lineH, thisLine, lineNumber As Integer
    Sheets("day").Select
    highUnit = 25
    'ap8 值班 az8 休 7~23
'    For i = 7 To 23 ' 值班 休 竖向处理
'        Range("ap" & i).Select
'        lineNumber = UBound(Split(Range("ap" & i).Value, Chr(10))) + 1
'        If lineNumber > 6 Then
'            'Range("ap" & i).Value = "略"
'            'Range("BL" & i).Value = 2
'        End If
'
'        Range("az" & i).Select
'        lineNumber = UBound(Split(Range("az" & i).Value, Chr(10))) + 1
'        If lineNumber > 6 Then
'            'Range("az" & i).Value = "略"
'            'Range("BL" & i).Value = 2
'        End If
'    Next
'
'    For i = 6 To 23 '纵向 从左至右判断每一列内单行人名数，按最大值取行高
'        thisLine = i + 5
'        If Range("BL" & thisLine).Value > 0 Then
'            lineH = Range("BL" & thisLine).Value * highUnit
'        Else
'            lineH = highUnit
'        End If
'        If lineH > 409 Then lineH = 409
'        Rows(thisLine).RowHeight = lineH
'    Next
    
    lineH = 25
    Rows("6:25").RowHeight = lineH
    For I = 6 To 26
        Range("D" & I).Select
        Rows(I).EntireRow.AutoFit
        'If Rows(I).RowHeight < 390 Then
        Rows(I).RowHeight = Rows(I).RowHeight + 10
        'If I = 32 Then Rows(I).RowHeight = 25
    Next
'    Rows("14").Hidden = True

    Call dayColumnsWidth

    
End Function

Function dayColumnsWidth()
    Dim lineH, thisLine, lineNumber, maxLen, thisLen As Integer
    Dim tmpNames As Variant
    maxLen = 0: thisLen = 0
    For I = 1 To 23 '1 = k 事假 11 '0值隐藏
        thisCol = GetColumnName(2 * I - 1 + 10)
        Range(thisCol & "24").Select
        preCol = GetColumnName(2 * I - 1 + 9)
        If Range(thisCol & "24").Value = 0 Then
            Debug.Print preCol & "-" & thisCol
            Columns(preCol).Hidden = True
            Columns(thisCol).Hidden = True
        Else
            For j = 6 To 23 '行数
                If Range(preCol & j).Value <> "" Then
                    tmpNames = Split(Range(preCol & j).Value, Chr(10))
                    For k = 0 To UBound(tmpNames)
                        thisLen = Len(tmpNames(k))
                        If maxLen < thisLen Then maxLen = thisLen
                    Next
                    'Columns(preCol).AutoFit
                End If
            Next
            If maxLen < 3 Then maxLen = 3
            Columns(preCol).ColumnWidth = maxLen * 1.6 + 15
            Columns(thisCol).ColumnWidth = 15
            maxLen = 0
        End If
    Next
'    Columns("H").ColumnWidth = 33
'    Columns("R").ColumnWidth = 33
'    Columns("AP").ColumnWidth = 33
'    Columns("AZ").ColumnWidth = 33
'    Columns("D:D").ColumnWidth = 38
    ''Columns("H").Hidden = True
    Columns("A").Hidden = True
'    If startLine > 6 Then Rows("40").Hidden = True
'    Rows("14").Hidden = True
End Function

Function WeekColumnsWidth()
    Dim lineH, thisLine, lineNumber, maxLen, thisLen As Integer
    Dim tmpNames As Variant
    maxLen = 0: thisLen = 0
    For I = 1 To 23 '1 = k 事假 11 '0值隐藏 16 旷工
        thisCol = GetColumnName(14 + I)
        Debug.Print "WeekColumnsWidth-" & thisCol
        Range(thisCol & "24").Select
        If Range(thisCol & "24").Value = 0 Then
            Columns(thisCol).Hidden = True
        Else
            For j = 32 To 49 '行数
                If Range(thisCol & j).Value <> "" Then
                    tmpNames = Split(Range(thisCol & j).Value, Chr(10))
                    For k = 0 To UBound(tmpNames)
                        thisLen = Len(tmpNames(k))
                        If maxLen < thisLen Then maxLen = thisLen
                    Next
                    'Columns(preCol).AutoFit
                End If
            Next
            If maxLen < 3 Then maxLen = 3
            Columns(thisCol).ColumnWidth = maxLen * 1.6 + 10
            maxLen = 0
        End If
    Next
    ''Columns("H").Hidden = True
    Columns("A").Hidden = True
'    If startLine > 6 Then Rows("40").Hidden = True
'    Rows("14").Hidden = True
End Function
Function lessNameHide()
    Dim thisCol As String
    For I = 15 To 37 ' E 事假 AK
        thisCol = GetColumnName(I)
        Range(thisCol & "24").Select
        If Range(thisCol & "24").Value = 0 Then
            Debug.Print thisCol
            Columns(thisCol).Hidden = True
        Else
            Columns(thisCol).AutoFit
        End If
    Next
End Function

Function Autofill()
        Rows("3:3").Select
        Range("AS3").Activate
        Selection.AutoFilter
        Range("V3:V4").Select
        ActiveSheet.Range("$A$3:$AR$176").AutoFilter Field:=22, Criteria1:="<>"
        Columns("V:V").ColumnWidth = 100
End Function

Function dayReport()
    Sheets("计算").Select
    totalLine = 176
    Dim columnNumber As Integer
    Dim posiPre As String
    Dim posi As String
    columnNumber = 7
    posiPre = GetColumnName(8 + 2 * (columnNumber - 1))
    posi = GetColumnName(GetColumnNum(posiPre) + 1)
    Debug.Print posiPre
    Debug.Print posi
    
    Dim lineNum, dataStartLine, tmpSum As Integer
    lineNum = 2:    dataStartLine = 347
    For I = 0 To 21
        If Range(posi & dataStartLine + I).Value > 0 And I <> 10 Then '离职：
            lineNum = lineNum + 1
            Range(posiPre & 318 + I).Value = lineNum & Chr(10) & Range(posiPre & 318 + I).Value
            tmpSum = tmpSum + Range(posi & dataStartLine + I).Value
        End If
    Next I
    For Line = 7 To totalLine + 6
        cellName = posiPre & Line '1号前面一列
'        If Line = 134 Then
'            Debug.Print
'        End If
        For Row = 1 To 1
            Call dealReport(posiPre, posi)
        Next Row
        If (Line Mod 50) = 0 Then
            Range(posiPre & "339").Select
            Delayms (2)
        End If
        ColumnName = ""
        cellName = ""
    Next Line
    Range(posiPre & "318").Select
    'ActiveWorkbook.Save
End Function


Function dealReport(str As String, posi As String) As String
    Debug.Print posi
    Dim posiRow As Integer
    posiRow = 318
    Range(cellName).Select
    Selection.Offset(0, 1).Select '右移1个单元格
    cellName = Selection.Address
    ColumnName = GetColumnName(Selection.Column)
    Debug.Print "上午--" & Range("C" & Selection.Row).Value & "--行号--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
    'If IsNumeric(Range(Selection.Address).Value) Then
    '    Range(posi & posiRow).Value = Range(posi & posiRow).Value & "，" & Range("C" & Selection.Row).Value & "-" & Range(Selection.Address).Value & "分钟" '取姓名及值
    '    Debug.Print Range(posi & posiRow).Value
    If Range(Selection.Address).Value = "迟到" Then
        Range(posi & posiRow).Value = Range(posi & posiRow).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
        Debug.Print Range(posi & posiRow).Value
    ElseIf Range(Selection.Address).Value = "值班旷工" Then
        Range(posi & posiRow + 20 + 1).Value = Range(posi & posiRow + 20 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
        Debug.Print Range(posi & posiRow + 20 + 1).Value
        Range(posi & posiRow + 20 + 1).Select
    ElseIf InStr(Range(Selection.Address).Value, "旷") Then
        If Range(Selection.Address).Value = "旷工" Then
            Range(posi & posiRow + 1 + 1).Value = Range(posi & posiRow + 1 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
        Else
            Range(posi & posiRow + 1 + 1).Value = Range(posi & posiRow + 1 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时" '取姓名及值
        End If
        Debug.Print Range(posi & posiRow + 1 + 1).Value
    ElseIf Range(Selection.Address).Value = "出差" Then
        Range(posi & posiRow + 2 + 1).Value = Range(posi & posiRow + 2 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
        Debug.Print Range(posi & posiRow + 2 + 1).Value
    ElseIf Range(Selection.Address).Value = "倒休" Then
        Range(posi & posiRow + 3 + 1).Value = Range(posi & posiRow + 3 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
        Debug.Print Range(posi & posiRow + 3 + 1).Value
    ElseIf Range(Selection.Address).Value = "年假" Then
        Range(posi & posiRow + 10 + 1).Value = Range(posi & posiRow + 10 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
        Debug.Print Range(posi & posiRow + 10 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "年") Then
        Range(posi & posiRow + 10 + 1).Value = Range(posi & posiRow + 10 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时" '取姓名及值
        Debug.Print Range(posi & posiRow + 10 + 1).Value
    ElseIf Range(Selection.Address).Value = "休" Then
        Range(posi & posiRow + 4 + 1).Value = Range(posi & posiRow + 4 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
        Debug.Print Range(posi & posiRow + 4 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "事") Then
        If Range(Selection.Address).Value = "事假" Then
            Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
        Else
            Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时" '取姓名及值
        End If
        Debug.Print Range(posi & posiRow + 5 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "病") Then
        If Range(Selection.Address).Value = "病假" Then
            Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
        Else
            Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时"  '取姓名及值
        End If
        Debug.Print Range(posi & posiRow + 6 + 1).Value
    ElseIf Range(Selection.Address).Value = "未签到" Then
        Range(posi & posiRow + 7 + 1).Value = Range(posi & posiRow + 7 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
        Debug.Print Range(posi & posiRow + 7 + 1).Value
    ElseIf Range(Selection.Address).Value = "补签到" Then
        Range(posi & posiRow + 14 + 1).Value = Range(posi & posiRow + 14 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
        Debug.Print Range(posi & posiRow + 14 + 1).Value
'    ElseIf Range(Selection.Address).Value = "补卡" Then
'        Range(posi & posiRow + 14 + 1).Value = Range(posi & posiRow + 14 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
        Debug.Print Range(posi & posiRow + 14 + 1).Value
    ElseIf Range(Selection.Address).Value = "入职" Then
        Range(posi & posiRow + 8 + 1).Value = Range(posi & posiRow + 8 + 1).Value & "，" & Range("C" & Selection.Row).Value    '取姓名
        Debug.Print Range(posi & posiRow + 8 + 1).Value
    ElseIf Range(Selection.Address).Value = "离职" Then
        Range(posi & posiRow + 9 + 1).Value = Range(posi & posiRow + 9 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
        Debug.Print Range(posi & posiRow + 9 + 1).Value
    ElseIf Range(Selection.Address).Value = "产假" Then
        Range(posi & posiRow + 12 + 1).Value = Range(posi & posiRow + 12 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
        Debug.Print Range(posi & posiRow + 12 + 1).Value
    ElseIf Range(Selection.Address).Value = "外出" Then
        Range(posi & posiRow + 13 + 1).Value = Range(posi & posiRow + 13 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
        Debug.Print Range(posi & posiRow + 13 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "休") Then '倒休1 2 3
        Range(posi & posiRow + 3 + 1).Value = Range(posi & posiRow + 3 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时" '取姓名及值
        Debug.Print Range(posi & posiRow + 3 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "补迟") Then '
        Range(posi & posiRow + 12 + 1).Value = Range(posi & posiRow + 12 + 1).Value & "，" & Range("C" & Selection.Row).Value    '取姓名
        Debug.Print Range(posi & posiRow + 12 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "补卡") Then '
        Range(posi & posiRow + 14 + 1).Value = Range(posi & posiRow + 14 + 1).Value & "，" & Range("C" & Selection.Row).Value    '取姓名
        Debug.Print Range(posi & posiRow + 14 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "丧假") Then '
        Range(posi & posiRow + 15 + 1).Value = Range(posi & posiRow + 15 + 1).Value & "，" & Range("C" & Selection.Row).Value    '取姓名
        Debug.Print Range(posi & posiRow + 15 + 1).Value
'    ElseIf InStr(Range(Selection.Address).Value, "调休") Then '
'        Range(posi & posiRow + 17 + 1).Value = Range(posi & posiRow + 17 + 1).Value & "，" & Range("C" & Selection.Row).Value    '取姓名
'        Debug.Print Range(posi & posiRow + 17 + 1).Value
    End If
    
    Selection.Offset(0, 1).Select '右移1个单元格
    cellName = Selection.Address
    Debug.Print "next cell " & cellName
    ColumnName = GetColumnName(Selection.Column)
    Debug.Print "下午--" & Range("C" & Selection.Row).Value & "--行号--" & ColumnName & Selection.Row & "===" & Range(Selection.Address).Value
    
    If IsNumeric(Range(Selection.Address).Value) Then
        If Not InStr(Range(posi & posiRow + 1).Value, Range("C" & Selection.Row).Value) Then
            Range(posi & posiRow + 1).Value = Range(posi & posiRow + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & -Range(Selection.Address).Value & "分钟"   '取姓名及值
            Debug.Print Range(posi & posiRow + 1).Value
            Range(posi & posiRow + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "值班旷工" Then
        If Not InStr(Range(posi & posiRow + 20 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 20 + 1).Value = Range(posi & posiRow + 20 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
            Debug.Print Range(posi & posiRow + 20 + 1).Value
            Range(posi & posiRow + 20 + 1).Select
        End If
    ElseIf InStr(Range(Selection.Address).Value, "旷") Then
        If Not InStr(Range(posi & posiRow + 1 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            If Range(Selection.Address).Value = "旷工" Then
                Range(posi & posiRow + 1 + 1).Value = Range(posi & posiRow + 1 + 1).Value & "，" & Range("C" & Selection.Row).Value    '取姓名
            Else
                Range(posi & posiRow + 1 + 1).Value = Range(posi & posiRow + 1 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时" '取姓名及值
            End If
            Debug.Print Range(posi & posiRow + 1).Value
            Range(posi & posiRow + 1 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "倒休" Then
        If Not InStr(Range(posi & posiRow + 3 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 3 + 1).Value = Range(posi & posiRow + 3 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
            Debug.Print Range(posi & posiRow + 3).Value
            Range(posi & posiRow + 3 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "年假" Then
        If Not InStr(Range(posi & posiRow + 10 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 10 + 1).Value = Range(posi & posiRow + 10 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
            Debug.Print Range(posi & posiRow + 10 + 1).Value
            Range(posi & posiRow + 10 + 1).Select
        End If
    ElseIf InStr(Range(Selection.Address).Value, "年") Then
        If Not InStr(Range(posi & posiRow + 10 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 10 + 1).Value = Range(posi & posiRow + 10 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1)  '取姓名及值
            Debug.Print Range(posi & posiRow + 10 + 1).Value
            Range(posi & posiRow + 10 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "休" Then
        If Not InStr(Range(posi & posiRow + 4 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 4 + 1).Value = Range(posi & posiRow + 4 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
            Debug.Print Range(posi & posiRow + 4 + 1).Value
            Range(posi & posiRow + 4 + 1).Select
        End If
'    ElseIf Range(Selection.Address).Value = "补卡" Then
'        If Not InStr(Range(posi & posiRow + 14 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
'            Range(posi & posiRow + 14 + 1).Value = Range(posi & posiRow + 14 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
'            Debug.Print Range(posi & posiRow + 14 + 1).Value
'            Range(posi & posiRow + 14 + 1).Select
'        End If
    ElseIf Range(Selection.Address).Value = "外出" Then '原补退
        If Not InStr(Range(posi & posiRow + 13 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 13 + 1).Value = Range(posi & posiRow + 13 + 1).Value & "，" & Range("C" & Selection.Row).Value    '取姓名
            Debug.Print Range(posi & posiRow + 13 + 1).Value
            Range(posi & posiRow + 13 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "未签退" Then
        If Not InStr(Range(posi & posiRow + 16 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 16 + 1).Value = Range(posi & posiRow + 16 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
            Debug.Print Range(posi & posiRow + 16 + 1).Value
            Range(posi & posiRow + 16 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "补签退" Then
        If Not InStr(Range(posi & posiRow + 17 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 17 + 1).Value = Range(posi & posiRow + 17 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
            Debug.Print Range(posi & posiRow + 17 + 1).Value
            Range(posi & posiRow + 17 + 1).Select
        End If
'    ElseIf Range(Selection.Address).Value = "调休" Then
'        If Not InStr(Range(posi & posiRow + 17 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
'            Range(posi & posiRow + 17 + 1).Value = Range(posi & posiRow + 17 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
'            Debug.Print Range(posi & posiRow + 17 + 1).Value
'            Range(posi & posiRow + 17 + 1).Select
'        End If
    ElseIf Range(Selection.Address).Value = "公假" Then
        If Not InStr(Range(posi & posiRow + 18 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 18 + 1).Value = Range(posi & posiRow + 18 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
            Debug.Print Range(posi & posiRow + 18 + 1).Value
            Range(posi & posiRow + 18 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "工伤假" Then
        If Not InStr(Range(posi & posiRow + 19 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 19 + 1).Value = Range(posi & posiRow + 19 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
            Debug.Print Range(posi & posiRow + 19 + 1).Value
            Range(posi & posiRow + 19 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "丧假" Then
        If Not InStr(Range(posi & posiRow + 15 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 15 + 1).Value = Range(posi & posiRow + 15 + 1).Value & "，" & Range("C" & Selection.Row).Value   '取姓名
            Debug.Print Range(posi & posiRow + 15 + 1).Value
            Range(posi & posiRow + 15 + 1).Select
        End If
    ElseIf InStr(Range(Selection.Address).Value, "事") Then
        If Not InStr(Range(posi & posiRow + 5 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时" '取姓名及值
            Range(posi & posiRow + 5 + 1).Select
'        Else
'            If Range(Selection.Address).Value = "事假" Then
'                Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
'            Else
'                Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时" '取姓名及值
'            End If
        End If
        Debug.Print Range(posi & posiRow + 5 + 1).Value
        
    ElseIf InStr(Range(Selection.Address).Value, "病") Then
        If Not InStr(Range(posi & posiRow + 6 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1)  '取姓名及值
            Debug.Print Range(posi & posiRow + 6 + 1).Value
            Range(posi & posiRow + 6 + 1).Select
'        Else
'            If Range(Selection.Address).Value = "病假" Then
'                Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "，" & Range("C" & Selection.Row).Value  '取姓名
'            Else
'                Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时"  '取姓名及值
'            End If
        End If
    ElseIf InStr(Range(Selection.Address).Value, "休") And Len(Range(Selection.Address).Value) > 1 Then
        If Not InStr(Range(posi & posiRow + 3 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 3 + 1).Value = Range(posi & posiRow + 3 + 1).Value & "，" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "小时" '取姓名及值
            Debug.Print Range(posi & posiRow + 3).Value
            Range(posi & posiRow + 3 + 1).Select
        End If
    End If
End Function

Function dealBlank() As String
    Sheets("计算").Select
    For I = 7 To 210
        Range("H" & I).Select
        For j = 1 To 50
            Selection.Offset(0, 1).Select '右移1个单元格
            Debug.Print Range("K8").Value
            cellName = Selection.Address
            ColumnName = GetColumnName(Selection.Column)
            Debug.Print "上午--" & Range("C" & Selection.Row).Value & "--行号--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
            If Trim(Range(Selection.Address).Value) = "" Or Trim(Range(Selection.Address).Value) = 0 Then
                Range(Selection.Address).Value = "休"
            End If
        Next j
    Next I
End Function

Function justSecquence()
    Dim tmpName As String
    
    For I = 7 To 176
        Sheets("计算").Select
        tmpName = Range("C" & I).Value
        Debug.Print tmpName
        Sheets("W1").Select
        If Range("E" & I).Value <> tmpName Then
            Cells.Find(What:=tmpName, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                , MatchByte:=False, SearchFormat:=False).Activate
            tmpRow = Selection.Row
            Debug.Print tmpName & ">姓名及行数>" & tmpRow
            Rows(tmpRow & ":" & tmpRow).Select
            Selection.Cut
            Rows(I & ":" & I).Select
            Selection.Insert Shift:=xlDown
        End If
        Delayms (1)
    Next
End Function

Function ll()
    Sheets("计算").Select
    Dim posi As String
    posi = "H"
    recordLine = 410
    LogNum = 1
    For Line = 7 To 176
        cellName = posi & Line '1号前面一列
        For Row = 1 To 1
            Call dealDate
        Next Row
        If Line = 46 Then LogNum = 1
        If vStr <> "" Then
            Range(posi & recordLine).Value = vStr
            vStr = ""
            recordLine = recordLine + 1
            LogNum = LogNum + 1
        End If
        cellName = ""
        If (Line Mod 50) = 0 Then
            Range(posi & recordLine).Select
            'ActiveWorkbook.Save
            Delayms (2)
        End If
        ColumnName = ""
        cellName = ""
        Range(posi & recordLine).Select
    Next Line
    Range(posi & "7").Select
    'ActiveWorkbook.Save
End Function

Function dealDate() As String
    Range(cellName).Select
    Selection.Offset(0, 1).Select '右移1个单元格
    cellName = Selection.Address
    ColumnName = GetColumnName(Selection.Column)
    Debug.Print "上午--" & Range("C" & Selection.Row).Value & "--行号--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
    If Range(Selection.Address).Value = "" Or Range(Selection.Address).Value = 0 Or InStr(Range(Selection.Address).Value, "星期") Then
        Range(Selection.Address).Value = "休"
    ElseIf Range(Selection.Address).Value = "迟到" Or (IsNumeric(Range(Selection.Address).Value) And Range(Selection.Address).Value > 0) Then
        'la (Selection.Address)
        If Range(Selection.Address).Value = "迟到" Then
            Range(ColumnName & "1").Select
            Selection.Copy
            Range(cellName).Select
            ActiveSheet.Paste
            'ActiveWorkbook.Save
        End If
        If Range(Selection.Address).Value >= 30 Then
            If vStr = "" Then
                     vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                            Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "迟到" & Range(Selection.Address).Value & "分钟"
                Debug.Print vStr
            Else
                vStr = vStr & "," & " " & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "迟到" & Range(Selection.Address).Value & "分钟"
                Debug.Print vStr
            End If
        End If
        Debug.Print Selection.Address & "迟到"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "缺卡" Or Range(Selection.Address).Value = "未签到" Then
        Range(Selection.Address).Value = "未签到"
        If vStr = "" Then
            vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "未签到"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "未签到"
        End If
        Debug.Print Selection.Address & "未签到"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "补签到" Then
        Range(Selection.Address).Value = "补签到"
        If vStr = "" Then
            vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "补签到"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "补签到"
        End If
        Debug.Print Selection.Address & "补签到"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "旷工" Then
        If vStr = "" Then
            vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                    Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & "旷工"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & "旷工"
        End If
        Debug.Print Selection.Address & "旷工"
        Debug.Print vStr
    End If
    
    Selection.Offset(0, 1).Select '右移1个单元格
    cellName = Selection.Address
    ColumnName = GetColumnName(Selection.Column)
    Debug.Print "下午--" & Range("C" & Selection.Row).Value & "--行号--" & ColumnName & Selection.Row & "===" & Range(Selection.Address).Value
    
    If Range(Selection.Address).Value = "" Or Range(Selection.Address).Value = 0 Or InStr(Range(Selection.Address).Value, "星期") Then
        Range(Selection.Address).Value = "休"
    ElseIf Range(Selection.Address).Value = "早退" Or (IsNumeric(Range(Selection.Address).Value) And Range(Selection.Address).Value < 0) Then
        If Range(Selection.Address).Value = "早退" Then 'le (Selection.Address)
            Range(ColumnName & "1").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range(cellName).Select
            ActiveSheet.Paste
            'ActiveWorkbook.Save
        End If
        
        If Range(Selection.Address).Value <= 30 Then
            If vStr = "" Then
              vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                     Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "早退" & -Range(Selection.Address).Value & "分钟"
                Debug.Print vStr
            Else
                vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "早退" & -Range(Selection.Address).Value & "分钟"
                Debug.Print vSt
            End If
        End If
    ElseIf Range(Selection.Address).Value = "缺卡" Or Range(Selection.Address).Value = "未签退" Then
        Range(Selection.Address).Value = "未签退"
        If vStr = "" Then
            vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "未签退"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "未签退"
        End If
        Debug.Print Selection.Address & "未签退"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "补签退" Then
        Range(Selection.Address).Value = "补签退"
        If vStr = "" Then
            vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "补签退"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "补签退"
        End If
        Debug.Print Selection.Address & "补签退"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "公假" Then
        Range(Selection.Address).Value = "公假"
        If vStr = "" Then
            vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "公假"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "公假"
        End If
        Debug.Print Selection.Address & "公假"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "工伤假" Then
        Range(Selection.Address).Value = "工伤假"
        If vStr = "" Then
            vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "工伤假"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "工伤假"
        End If
        Debug.Print Selection.Address & "工伤假"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "尚未打卡" Then
        Range(Selection.Address).Value = "正常"
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        If vStr = "" Then
            vStr = LogNum & "、 " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "尚未打卡"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "尚未打卡"
        End If
        Debug.Print Selection.Address & "尚未打卡"
        Debug.Print vStr
    End If
End Function

Function dayWeekReport()
    Sheets("计算").Select
    Dim posiPreLeft  As String, posiPreRight As String
    Dim posi, posiCounter As String
    hasChange = False
    cd = 0: zt = 0: kg = 0: cc = 0: dx = 0: zcx = 0: sj = 0: bj = 0: qk = 0: rz = 0: lz = 0: nx = 0: bc = 0: bt = 0: bk = 0: wqd = 0: wqt = 0: bqd = 0: bqt = 0  '迟到、早退、旷工、出差、倒休、事假、病假、缺卡、补迟、补退、补卡
    cds = 0: zts = 0: kgs = 0: ccs = 0: dxs = 0: zcxs = 0: sjs = 0: bjs = 0: qks = 0: rzs = 0: lzs = 0: nxs = 0: bcs = 0: bts = 0: bks = 0: wqds = 0: wqts = 0: bqds = 0: bqts = 0 '迟到、早退、旷工、出差、倒休、事假、病假、缺卡、补迟、补退、补卡
    cdStr = "":  ztStr = "":  kgStr = "":  ccStr = "":  dxStr = "": zcxStr = ""
    sjStr = "": bjStr = "": qkStr = "": rzStr = "": lzStr = "": nxStr = "": bcStr = "": btStr = "": bkStr = "": wqdStr = 0: wqtStr = 0: bqdStr = 0: bqtStr = 0
    
    posiPreLeft = "G" '周第一天开始前2列'-------------------------------------------
    
    posiName = GetColumnName(GetColumnNum(posiPreLeft) + 5) '第四列为人名
    posiCounter = GetColumnName(GetColumnNum(posiPreLeft) + 4) '第三列为人数
    Debug.Print posiPreLeft
    Debug.Print posiName
    Debug.Print posiCounter
    
    For dayN = 1 To 7
        'For Line = 117 To 117 '
        For Line = 7 To 176 '-------------------------------------------
            cellName = GetColumnName(GetColumnNum(posiPreLeft) + (dayN - 1) * 2) & Line '1号前面一列
            Debug.Print cellName
            For Row = 1 To 1 '周天数
                Call dealWeekReportLeftSimple(posiPreLeft, "" & posiName)
            Next Row
            
            If (Line Mod 50) = 0 Then
                Range(posiPreLeft & "380").Select
                Delayms (2)
            End If
            
            ColumnName = ""
            cellName = ""
            'If hasChange Then ActiveWorkbook.Save
            hasChange = False
        Next Line
        Range(posiPreLeft & "380").Select
        
        posiPreRight = "H" '周天开始前1列'-------------------------------------------
        'For Line = 117 To 117 '
        For Line = 7 To 176
            cellName = GetColumnName(GetColumnNum(posiPreRight) + (dayN - 1) * 2) & Line '1号前面一列
            'Debug.Print cellName
            For Row = 1 To 1 '周天数
                Call dealWeekReportRightSimple(posiPreRight, "" & posiName)
            Next Row
            
            If (Line Mod 50) = 0 Then
                Range(posiPreRight & "390").Select
                Delayms (2)
            End If
            
            ColumnName = ""
            Range(posiPreRight & "380").Select
            'If hasChange Then ActiveWorkbook.Save
            hasChange = False
            cellName = ""
        Next Line
        Debug.Print "cdStr=" & cdStr
        Debug.Print "ztStr=" & ztStr
        Debug.Print "kgStr=" & kgStr
        Debug.Print "ccStr=" & ccStr
        Debug.Print "dxStr=" & dxStr
        Debug.Print "sjStr=" & sjStr
        Debug.Print "bjStr=" & bjStr
        Debug.Print "wqdStr=" & wqdStr
        Debug.Print "wqtStr=" & wqtStr
        Debug.Print "nxStr=" & nxStr
        Debug.Print "bcStr=" & bcStr
        Debug.Print "btStr=" & btStr
        Debug.Print "bqdStr=" & bqdStr
        Debug.Print "bqtStr=" & bqtStr
    Next dayN
    
     Dim lineNum  As Integer, startLine As Integer, dateColume As String
     
     lineNum = 2: startLine = 380: dateColume = "I" '-------------------------------------------"
     
     For I = startLine To startLine + 18

             Select Case I
                 Case startLine
                 If cds > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = cd & "人"
                     Range(posiName & I).Value = cdStr
                     If Range(dateColume & I).Value <> cds Then
                        Range(dateColume & I).Value = cds
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                    End If
                 Case startLine + 1
                If zts > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = cd & "人"
                     Range(posiCounter & I).Value = zt & "人"
                     Range(posiName & I).Value = ztStr
                     If Range(dateColume & I).Value <> zts Then
                        Range(dateColume & I).Value = zts
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                    End If
                 Case startLine + 2
                If kgs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = kg & "人"
                     Range(posiName & I).Value = kgStr
                     If Range(dateColume & I).Value <> kgs Then
                        Range(dateColume & I).Value = kgs
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 3
                If ccs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = cc & "人"
                     Range(posiName & I).Value = ccStr
                     If Range(dateColume & I).Value <> ccs Then
                        Range(dateColume & I).Value = ccs / 2
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 4
                If dxs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = dx & "人"
                     Range(posiName & I).Value = dxStr
                     If Range(dateColume & I).Value <> dxs Then
                        Range(dateColume & I).Value = dxs
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 5
                If zcx > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = zcx & "人"
                     Range(dateColume & I).Value = zcx
                     End If
                 Case startLine + 6
                If sjs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = sj & "人"
                     Range(posiName & I).Value = sjStr
                     If Range(dateColume & I).Value <> sjs Then
                        Range(dateColume & I).Value = sjs
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 7
                If bjs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = bj & "人"
                     Range(posiName & I).Value = bjStr
                     If Range(dateColume & I).Value <> bjs Then
                        Range(dateColume & I).Value = bjs
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 8
                If wqds > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = wqd & "人"
                     Range(posiName & I).Value = wqdStr
                     If Range(dateColume & I).Value <> wqds Then
                        Range(dateColume & I).Value = wqds
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 9
                If rz > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = rz & "人"
                     End If
                 Case startLine + 10
                If lz > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = lz & "人"
                     End If
                 Case startLine + 11
                If nxs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = nx & "人"
                     Range(posiName & I).Value = nxStr
                     If Range(dateColume & I).Value <> nxs Then
                        Range(dateColume & I).Value = nxs
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 13 'bc
                If bcs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = bc & "人"
                     Range(posiName & I).Value = bcStr
                     If Range(dateColume & I).Value <> bcs Then
                        Range(dateColume & I).Value = bcs
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 14 'bt
                If bts > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = bt & "人"
                     Range(posiName & I).Value = btStr
                     If Range(dateColume & I).Value <> bts Then
                        Range(dateColume & I).Value = bts
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 15 'bqd
                If bqd > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = bqd & "人"
                     Range(posiName & I).Value = bqdStr
                     If Range(dateColume & I).Value <> bqds Then
                        Range(dateColume & I).Value = bqds
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 17 'wqt
                If wqts > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = wqt & "人"
                     Range(posiName & I).Value = wqtStr
                     If Range(dateColume & I).Value <> wqts Then
                        Range(dateColume & I).Value = wqts
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
                 Case startLine + 18 'bqt
                If bqts > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = bqt & "人"
                     Range(posiName & I).Value = bqtStr
                     If Range(dateColume & I).Value <> bqts Then
                        Range(dateColume & I).Value = bqts
                        Range(dateColume & I).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 14351096
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                     End If
                     End If
             End Select
     Next I
    cd = 0: zt = 0: kg = 0: cc = 0: dx = 0: zcx = 0: sj = 0: bj = 0: qk = 0: rz = 0: lz = 0: nx = 0: bc = 0: bt = 0: bqd = 0: bqt = 0: wqd = 0: wqt = 0 '迟到、早退、旷工、出差、倒休、事假、病假、缺卡、补迟、补退、补卡
    cds = 0: zts = 0: kgs = 0: ccs = 0: dxs = 0: zcxs = 0: sjs = 0: bjs = 0: qks = 0: rzs = 0: lzs = 0: nxs = 0: bcs = 0: bts = 0: bqds = 0: bqts = 0: wqds = 0: wqts = 0 '迟到、早退、旷工、出差、倒休、事假、病假、缺卡、补迟、补退、补卡
    cdStr = "":  ztStr = "":  kgStr = "":  ccStr = "":  dxStr = "": zcxStr = ""
    sjStr = "": bjStr = "": qkStr = "": rzStr = "": lzStr = "": nxStr = "": bcStr = "": btStr = "": bqdStr = "": wqdStr = 0: wqtStr = 0: bqdStr = 0: bqtStr = 0
 
    'lineNum = 2:    startLine = 344
    'For i = 0 To 12
    '    If Range("I" & startLine + i).Value > 0 Then
    '        lineNum = lineNum + 1
    '        Range(posiPreRight & startLine + i).Value = lineNum & chr(10) & Range(posiPreRight & startLine + i).Value
    '    End If
    'Next i
    Range(posiPreRight & "400").Select
    'ActiveWorkbook.Save
End Function


Function hide12()
    Sheets("计算").Select
    Rows("240:307").Select
    Selection.EntireRow.Hidden = False
    For I = 240 To 307
        If Range("H" & I).Value = 0 Then
            Rows(I & ":" & I).Select
            Selection.EntireRow.Hidden = True
        End If
    Next
'
'    For i = 94 To 126
'        If Range("H" & i).Value = 0 Then
'            Rows(i & ":" & i).Select
'            Selection.EntireRow.Hidden = True
'        End If
'    Next
    ActiveWorkbook.Save
End Function

Function hide1()
    Sheets("公示1").Select
    Rows("51:121").Select
    Selection.EntireRow.Hidden = False
    For I = 51 To 81
        If Range("H" & I).Value = 0 Then
            Rows(I & ":" & I).Select
            Selection.EntireRow.Hidden = True
        End If
    Next
    
    For I = 90 To 120
        If Range("H" & I).Value = 0 Then
            Rows(I & ":" & I).Select
            Selection.EntireRow.Hidden = True
        End If
    Next
    ActiveWorkbook.Save
End Function

Function dealWeekReportLeftSimple(str As String, posi As String) As String
    'Debug.Print posi
    Dim posiRow As Integer
    Dim findPosi As Integer
    Dim t1 As String, t2 As String
    posiRow = 380
    Range(cellName).Select
    Selection.Offset(0, 2).Select '右移2个单元格
    cellName = Selection.Address
    ColumnName = GetColumnName(Selection.Column)
    'Debug.Print "上午--" & Range("C" & Selection.Row).Value & "--行号--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
    If Range(Selection.Address).Value = "迟到" Then
        findPosi = InStr(cdStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        If findPosi <= 0 Then
            cdStr = cdStr & "，" & Range("C" & Selection.Row).Value '取姓名 第一次
            cd = cd + 1
        Else
            If Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(cdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                cdStr = Replace(cdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(cdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                cdStr = Replace(cdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then   '第二次
                cdStr = Replace(cdStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
            '1.我有字符串 Ilovevba,我如何提取 love = Mid("Lloveba", 2, 4) 开始，长度
        End If
        cds = cds + 1
        Debug.Print "迟到=" & cdStr
    ElseIf Range(Selection.Address).Value = "加班" Then '补迟
        findPosi = InStr(bcStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print bcStr
        If findPosi <= 0 Then
            'Debug.Print bcStr
            'Debug.Print Range("C" & Selection.Row).Value
            bcStr = bcStr & "，" & Range("C" & Selection.Row).Value '取姓名 第一次
            bc = bc + 1
        Else
            If Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(bcStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bcStr = Replace(bcStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(bcStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bcStr = Replace(bcStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                bcStr = Replace(bcStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "补迟=" & bcStr
        bcs = bcs + 1
    ElseIf Range(Selection.Address).Value = "补卡" Then
        findPosi = InStr(bqdStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print bkStr
        If findPosi <= 0 Then
            'Debug.Print bkStr
            'Debug.Print Range("C" & Selection.Row).Value
            bqdStr = bqdStr & "，" & Range("C" & Selection.Row).Value '取姓名 第一次
            bqd = bqd + 1
        Else
            If Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(bqdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bqdStr = Replace(bqdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(bqdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bqdStr = Replace(bqdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                bqdStr = Replace(bqdStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "补签到=" & bqdStr
        bqds = bqds + 1
    ElseIf Range(Selection.Address).Value = "外出" Then '补退
        findPosi = InStr(btStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            btStr = btStr & "，" & Range("C" & Selection.Row).Value '取姓名
            Range(posi & posiRow + 8).Select
            bt = bt + 1
        Else
            If Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(btStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                btStr = Replace(btStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(btStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                btStr = Replace(btStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                btStr = Replace(btStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "补退=" & btStr
        bts = bts + 1
    ElseIf InStr(Range(Selection.Address).Value, "旷") Then
        findPosi = InStr(kgStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print kgStr
        If findPosi <= 0 Then
            'Debug.Print kgStr
            'Debug.Print Range("C" & Selection.Row).Value
            kgStr = kgStr & "，" & Range("C" & Selection.Row).Value '取姓名 第一次
            kg = kg + 1
        Else
            If Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(kgStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                kgStr = Replace(kgStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(kgStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                kgStr = Replace(kgStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                kgStr = Replace(kgStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "旷工=" & kgStr
        kgs = kgs + 1
    ElseIf Range(Selection.Address).Value = "出差" Then
        'Debug.Print "出差=" & ccStr
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print Len(Range("C" & Selection.Row).Value)
        findPosi = InStr(ccStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            ccStr = ccStr & "，" & Range("C" & Selection.Row).Value '取姓名
            cc = cc + 1
        Else
            If Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(ccStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                ccStr = Replace(ccStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(ccStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                ccStr = Replace(ccStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                ccStr = Replace(ccStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "出差=" & ccStr
        ccs = ccs + 1
        Range(posi & posiRow + 3).Select
    ElseIf Range(Selection.Address).Value = "倒休" Then
        findPosi = InStr(dxStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            dxStr = dxStr & "，" & Range("C" & Selection.Row).Value '取姓名
            dx = dx + 1
        Else
            If Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                dxStr = Replace(dxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "倒休=" & dxStr
        dxs = dxs + 1
    ElseIf InStr(Range(Selection.Address).Value, "年") Then
        findPosi = InStr(nxStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            nxStr = nxStr & "，" & Range("C" & Selection.Row).Value '取姓名
            nx = nx + 1
        Else
            If Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(nxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                nxStr = Replace(nxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(nxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                nxStr = Replace(nxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                nxStr = Replace(nxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "年假=" & nxStr
        nxs = nxs + 1
    ElseIf Range(Selection.Address).Value = "休" Then
        findPosi = InStr(Range(posi & posiRow + 5).Value, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            'Range(posi & posiRow + 4 + 1).Value = Range(posi & posiRow + 4 + 1).Value & "，" & Range("C" & Selection.Row).Value '取姓名
            Debug.Print Range(posi & posiRow + 5).Value
            zcx = zcx + 1
        End If
    ElseIf InStr(Range(Selection.Address).Value, "事") Then
        findPosi = InStr(sjStr, Range("C" & Selection.Row).Value)
        If InStr(sjStr, Range("C" & Selection.Row).Value) <= 0 Then
            sjStr = sjStr & "，" & Range("C" & Selection.Row).Value '取姓名
            sj = sj + 1
        Else
            If Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(sjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                sjStr = Replace(sjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(sjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                sjStr = Replace(sjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                sjStr = Replace(sjStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "事假=" & sjStr
        sjs = sjs + 1
    ElseIf InStr(Range(Selection.Address).Value, "病") Then
        findPosi = InStr(bjStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            bjStr = bjStr & "，" & Range("C" & Selection.Row).Value '取姓名
            bj = bj + 1
        Else
            If Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(bjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bjStr = Replace(bjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(bjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bjStr = Replace(bjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                bjStr = Replace(bjStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "病假=" & bjStr
        bjs = bjs + 1
    ElseIf Range(Selection.Address).Value = "未签到" Then
        findPosi = InStr(qkStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            wqdStr = wqdStr & "，" & Range("C" & Selection.Row).Value '取姓名
            wqd = wqd + 1
        Else
            If Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(wqdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                wqdStr = Replace(wqdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(wqdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                wqdStr = Replace(wqdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                wqdStr = Replace(wqdStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "未签到=" & wqdStr
        wqds = wqds + 1
    ElseIf Range(Selection.Address).Value = "入职" Then
        If InStr(Range(posi & posiRow + 8 + 1).Value, Range("C" & Selection.Row).Value) <= 0 Then
            Range(posi & posiRow + 8 + 1).Value = Range(posi & posiRow + 8 + 1).Value & "，" & Range("C" & Selection.Row).Value '取姓名
            Debug.Print Range(posi & posiRow + 8 + 1).Value
            rz = rz + 1
        Else
        
        End If
    ElseIf Range(Selection.Address).Value = "离职" Then
        If InStr(Range(posi & posiRow + 9 + 1).Value, Range("C" & Selection.Row).Value) <= 0 Then
            Range(posi & posiRow + 9 + 1).Value = Range(posi & posiRow + 9 + 1).Value & "，" & Range("C" & Selection.Row).Value '取姓名
            Debug.Print Range(posi & posiRow + 9 + 1).Value
            lz = lz + 1
        Else
        
        End If
    ElseIf InStr(Range(Selection.Address).Value, "休") Then '倒休1 2 3
        findPosi = InStr(dxStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            dxStr = dxStr & "，" & Range("C" & Selection.Row).Value '取姓名
            dx = dx + 1
        Else
            If Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                ''Debug.Print t1
                ''Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                dxStr = Replace(dxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "倒休=" & dxStr
        dxs = dxs + 1
    End If
    findPosi = 0
    t1 = ""
    t2 = ""
End Function

Function dealWeekReportRightSimple(str As String, posi As String) As String
    'Debug.Print posi
    Dim posiRow As Integer
    Dim replaceTmp As String
    Dim findPosi As Integer
    Dim t1 As String, t2 As String
    posiRow = 380
    Range(cellName).Select
    Selection.Offset(0, 2).Select '右移2个单元格
    cellName = Selection.Address
    'Debug.Print "next cell " & cellName
    ColumnName = GetColumnName(Selection.Column)
    'Debug.Print "下午--" & Range("C" & Selection.Row).Value & "--行号--" & ColumnName & Selection.Row & "===" & Range(Selection.Address).Value
    
    If Range(Selection.Address).Value = "早退" Then
        findPosi = InStr(ztStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            ztStr = ztStr & "，" & Range("C" & Selection.Row).Value '取姓名
            Range(posi & posiRow + 1).Select
            zt = zt + 1
        Else
            If Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(ztStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                ztStr = Replace(ztStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(ztStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                ztStr = Replace(ztStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                ztStr = Replace(ztStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "早退=" & ztStr
        zts = zts + 1
    ElseIf InStr(Range(Selection.Address).Value, "旷") Then
        findPosi = InStr(kgStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            kgStr = kgStr & "，" & Range("C" & Selection.Row).Value '取姓名 第一次
            kg = kg + 1
        Else
            If Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(kgStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                t2 = Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                kgStr = Replace(kgStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(kgStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                t2 = Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                kgStr = Replace(kgStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                kgStr = Replace(kgStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "旷工=" & kgStr
        kgs = kgs + 1
    ElseIf Range(Selection.Address).Value = "倒休" Then
        Selection.Offset(0, -1).Select '往回看上午是否也倒休
        If Range(Selection.Address).Value <> "倒休" Then
            findPosi = InStr(dxStr, Range("C" & Selection.Row).Value)
            If findPosi <= 0 Then
                dxStr = dxStr & "，" & Range("C" & Selection.Row).Value '取姓名
                Range(posi & posiRow + 4).Select
                dx = dx + 1
            Else
                If Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                    t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                    'Debug.Print t1
                    'Debug.Print Len(Range("C" & Selection.Row).Value)
                    t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                    'Debug.Print t2
                    dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
                ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                    t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                    'Debug.Print t1
                    'Debug.Print Len(Range("C" & Selection.Row).Value)
                    t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                    'Debug.Print t2
                    dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
                ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                    dxStr = Replace(dxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
                End If
            End If
            Debug.Print "倒休=" & dxStr
            dxs = dxs + 1
        End If
        Selection.Offset(0, 1).Select '移回来
    ElseIf Range(Selection.Address).Value = "休" Then
        Selection.Offset(0, -1).Select '往回看上午是否也年休
        If Range(Selection.Address).Value <> "休" Then
            findPosi = InStr(Range(posi & posiRow + 5).Value, Range("C" & Selection.Row).Value)
            If findPosi <= 0 Then
                'Range(posi & posiRow + 4 + 1).Value = Range(posi & posiRow + 4 + 1).Value & "，" & Range("C" & Selection.Row).Value '取姓名
                Debug.Print Range(posi & posiRow + 5).Value
                zcx = zcx + 1
            End If
        End If
        Selection.Offset(0, 1).Select '移回来
    ElseIf Range(Selection.Address).Value = "年假" Then
        Selection.Offset(0, -1).Select '往回看上午是否也年休
        If Range(Selection.Address).Value <> "年休" Then
            findPosi = InStr(nxStr, Range("C" & Selection.Row).Value)
            If findPosi <= 0 Then
                nxStr = nxStr & "，" & Range("C" & Selection.Row).Value '取姓名
                Range(posi & posiRow + 11).Select
                nx = nx + 1
            Else
                If Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                    t1 = Mid(nxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                    'Debug.Print t1
                    'Debug.Print Len(Range("C" & Selection.Row).Value)
                    t2 = Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                    'Debug.Print t2
                    nxStr = Replace(nxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
                ElseIf Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                    t1 = Mid(nxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                    'Debug.Print t1
                    'Debug.Print Len(Range("C" & Selection.Row).Value)
                    t2 = Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                    'Debug.Print t2
                    nxStr = Replace(nxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
                ElseIf Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                    nxStr = Replace(nxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
                End If
            End If
            Debug.Print "年假=" & nxStr
            nxs = nxs + 1
        End If
        Selection.Offset(0, 1).Select '移回来
    ElseIf Range(Selection.Address).Value = "未签退" Then
        findPosi = InStr(wqtStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            wqtStr = wqtStr & "，" & Range("C" & Selection.Row).Value '取姓名
            Range(posi & posiRow + 8).Select
            wqt = wqt + 1
        Else
            If Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(wqtStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                wqtStr = Replace(wqtStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(wqtStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                wqtStr = Replace(wqtStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                wqtStr = Replace(wqtStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "未签退=" & wqtStr
        wqts = wqts + 1
    ElseIf Range(Selection.Address).Value = "加班" Then '补迟
        findPosi = InStr(bcStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print bcStr
        If findPosi <= 0 Then
            'Debug.Print bcStr
            'Debug.Print Range("C" & Selection.Row).Value
            bcStr = bcStr & "，" & Range("C" & Selection.Row).Value '取姓名 第一次
            bc = bc + 1
        Else
            If Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(bcStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bcStr = Replace(bcStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(bcStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bcStr = Replace(bcStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                bcStr = Replace(bcStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "补迟=" & bcStr
        bcs = bcs + 1
    ElseIf Range(Selection.Address).Value = "补卡" Then
        findPosi = InStr(bqtStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            bqtStr = bqtStr & "，" & Range("C" & Selection.Row).Value '取姓名
            Range(posi & posiRow + 8).Select
            bqt = bqt + 1
        Else
            If Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(bqtStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bqtStr = Replace(bqtStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(bqtStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bqtStr = Replace(bqtStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                bqtStr = Replace(bqtStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "补签退=" & bqtStr
        bqts = bqts + 1
    ElseIf Range(Selection.Address).Value = "外出" Then '补退
        findPosi = InStr(btStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            btStr = btStr & "，" & Range("C" & Selection.Row).Value '取姓名
            Range(posi & posiRow + 8).Select
            bt = bt + 1
        Else
            If Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(btStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                btStr = Replace(btStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(btStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                btStr = Replace(btStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                btStr = Replace(btStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "补退=" & btStr
        bts = bts + 1
    ElseIf InStr(Range(Selection.Address).Value, "事") Then
        findPosi = InStr(sjStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            sjStr = sjStr & "，" & Range("C" & Selection.Row).Value '取姓名
            Range(posi & posiRow + 6).Select
            sj = sj + 1
        Else
            If Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(sjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                sjStr = Replace(sjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(sjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                sjStr = Replace(sjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                sjStr = Replace(sjStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "事假=" & sjStr
        sjs = sjs + 1
    ElseIf InStr(Range(Selection.Address).Value, "病") Then
        findPosi = InStr(bjStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            bjStr = bjStr & "，" & Range("C" & Selection.Row).Value '取姓名
            Range(posi & posiRow + 7).Select
            bj = bj + 1
        Else
            If Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(bjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bjStr = Replace(bjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(bjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bjStr = Replace(bjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                bjStr = Replace(bjStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "病假=" & bjStr
        bjs = bjs + 1
    ElseIf InStr(Range(Selection.Address).Value, "休") And Len(Range(Selection.Address).Value) > 0 Then
        findPosi = InStr(dxStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            dxStr = dxStr & "，" & Range("C" & Selection.Row).Value '取姓名
            Range(posi & posiRow + 4).Select
            dx = dx + 1
        Else
            If Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "次" Then '两位数
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "次" Then '个位数
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "，" Or Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '第二次
                dxStr = Replace(dxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2次")
            End If
        End If
        Debug.Print "倒休=" & dxStr
        dxs = dxs + 1
    End If
    findPosi = 0
    t1 = ""
    t2 = ""
End Function


Function pdf()
    Dim tempName As String
    'Sheets("计算-数值").Select
    'Sheets("2月").Select
    For I = 1 To 18
        tempName = Range("C" & 216 + I).Value
        Range("C202").Value = tempName
        ActiveSheet.Range("$A$6:$BR$201").AutoFilter Field:=3, Criteria1:=tempName
        ChDir "D:\调度\团队号\202402"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "D:\调度\团队号\202402\" & I & " " & tempName & "-2024年2月考勤统计.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False
    Next
End Function

Function pdfG()
    Dim tempName As String
    For I = 1 To 318
        tempName = Sheets("计算").Range("D" & I + 6).Value
        ActiveSheet.Range("$A$6:$BR$304").AutoFilter Field:=4, Criteria1:=tempName
        ChDir "D:\调度\团队号\202307-"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "D:\调度\团队号\202307-\" & I & " " & tempName & "-2022年1月考勤统计.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False
    Next
End Function

Function shiftMark()
    Sheets("计算").Select
    Dim ShiftLine(), markColumnName()
    ShiftLine = Array(5, 6, 12, 13, 19, 20, 26, 27, 33, 34)
    markColumnName = Array("K", "M", "Y", "AA", "AM", "AO", "BA", "BC", "BO", "BQ")
    Dim thisName, tempName, nameColumnName As String
    Dim nameColumnStart As Integer
    nameColumnStart = GetColumnNum("AI") '------HK name start column
    For I = 0 To 9 '值班天数
        'For j = 0 To 16 '名字列 h to X
        For j = 0 To 3 '名字列 AI to AL
            nameColumnName = GetColumnName(nameColumnStart + j)
            tempName = Sheets("hk").Range(nameColumnName & ShiftLine(I)).Value
            Debug.Print tempName
            For rowLine = 11 To 100
                thisName = Range("C" & rowLine)
                If thisName = tempName Then
                    markColor (markColumnName(I) & rowLine)
                    thisName = "": tempName = ""
                    Exit For
                End If
            Next
        Next
    Next
'    ActiveWorkbook.Save
End Function
Function markColor(cellName)
    Sheets("计算").Select
    Range(cellName).Select
    If Selection.Interior.Color <> 5287936 Then
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5287936
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Offset(0, 1).Select '右移1个单元格
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5287936
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End Function


Function formatResult()
            
    Sheets("1月").Select
    totalLine = 176
    
    Dim formatColumnStart, formatColumnEnd, cols As Integer
    Dim tempNum As Double
    formatColumnStart = GetColumnNum("AQ") '实出勤前一列 GetColumnNum("BH") '1-15 本月打卡累计
    formatColumnEnd = GetColumnNum("BX")
    cols = formatColumnEnd - formatColumnStart
    For Line = 7 To totalLine
        cellName = "AR" & Line
        'Debug.Print (cellName)
        For col = 1 To cols '
            Range(cellName).Select
            Selection.Offset(0, 1).Select '右移1个单元格
            cellName = Selection.Address
            If IsNumeric(Selection.Value) Then
                tempNum = Selection.Value
                Debug.Print tempNum
                If Int(tempNum) = tempNum Then
                    Selection.NumberFormatLocal = "0_);"
                ElseIf Int(tempNum * 10) = tempNum * 10 Then
                    Selection.NumberFormatLocal = "0.0_);"
                ElseIf Int(tempNum * 100) = tempNum * 100 Then
                    Selection.NumberFormatLocal = "0.00_);"
                ElseIf Int(tempNum * 1000) = tempNum * 1000 Then
                    Selection.NumberFormatLocal = "0.000_);"
                ElseIf Int(tempNum * 10000) = tempNum * 10000 Then
                    Selection.NumberFormatLocal = "0.0000_);"
                End If
            End If
        Next
    Next Line
    'ActiveWorkbook.Save
End Function

Function formatResult2()
            
    Sheets("mo").Select
    totalLine = 218
    
    Dim formatColumnStart, formatColumnEnd, cols As Integer
    Dim tempNum As Double
    formatColumnStart = GetColumnNum("AL") '实出勤前一列 GetColumnNum("BH") '1-15 本月打卡累计
    formatColumnEnd = GetColumnNum("BL")
    cols = formatColumnEnd - formatColumnStart - 2
    For Line = 6 To totalLine
        cellName = "AL" & Line
        'Debug.Print (cellName)
        If Range("F" & Line).Value = "是" Then
            For col = 1 To cols '
                Range(cellName).Select
                If IsNumeric(Selection.Value) Then
                    Selection.Offset(0, 1).Select '右移1个单元格
                    cellName = Selection.Address
                    tempNum = Selection.Value
                    Debug.Print tempNum
                    If Int(tempNum) = tempNum Then
                        Selection.NumberFormatLocal = "0_);"
                    ElseIf Int(tempNum * 10) = tempNum * 10 Then
                        Selection.NumberFormatLocal = "0.0_);"
                    ElseIf Int(tempNum * 100) = tempNum * 100 Then
                        Selection.NumberFormatLocal = "0.00_);"
                    ElseIf Int(tempNum * 1000) = tempNum * 1000 Then
                        Selection.NumberFormatLocal = "0.000_);"
                    ElseIf Int(tempNum * 10000) = tempNum * 10000 Then
                        Selection.NumberFormatLocal = "0.0000_);"
                    End If
                End If
            Next
        End If
    Next Line
    'ActiveWorkbook.Save
End Function

Function replaceMonth()
    Cells.Replace What:="倒休1h", Replacement:="休1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="倒休2h", Replacement:="休2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="倒休3h", Replacement:="休3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="倒休4h", Replacement:="休4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="倒休5h", Replacement:="休5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="倒休6h", Replacement:="休6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="倒休7h", Replacement:="休7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值迟", Replacement:="迟", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值旷", Replacement:="旷", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值未", Replacement:="未", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="迟未", Replacement:="未", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="调休1h", Replacement:="休1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="调休2h", Replacement:="休2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="调休3h", Replacement:="休3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="调休4h", Replacement:="休4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="调休5h", Replacement:="休5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="调休6h", Replacement:="休6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="调休7h", Replacement:="休7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事假1h", Replacement:="事1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事假2h", Replacement:="事2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事假3h", Replacement:="事3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事假4h", Replacement:="事4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事假5h", Replacement:="事5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事假6h", Replacement:="事6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事假7h", Replacement:="事7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工1h", Replacement:="旷1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工2h", Replacement:="旷2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工3h", Replacement:="旷3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工4h", Replacement:="旷4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工5h", Replacement:="旷5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工6h", Replacement:="旷6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工7h", Replacement:="旷7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="1h", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="2h", Replacement:="2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="3h", Replacement:="3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="4h", Replacement:="4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="5h", Replacement:="5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="6h", Replacement:="6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="7h", Replacement:="7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病假1h", Replacement:="病1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病假2h", Replacement:="病2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病假3h", Replacement:="病3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病假4h", Replacement:="病4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病假5h", Replacement:="病5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病假6h", Replacement:="病6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病假7h", Replacement:="病7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="年假1h", Replacement:="年1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="年假2h", Replacement:="年2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="年假3h", Replacement:="年3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="年假4h", Replacement:="年4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="年假5h", Replacement:="年5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="年假6h", Replacement:="年6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="年假7h", Replacement:="年7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值班1h", Replacement:="值1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值班2h", Replacement:="值2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值班3h", Replacement:="值3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值班4h", Replacement:="值4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值班5h", Replacement:="值5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值班6h", Replacement:="值6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值班7h", Replacement:="值7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'    Cells.Replace What:="迟到", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'    Cells.Replace What:="早退", Replacement:="-", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
    Cells.Replace What:="m", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="SU(", Replacement:="SUM(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="出差值班", Replacement:="出差", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Function

Function replaceDayHour()
    Cells.Replace What:="1小时", Replacement:="0.125天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="2小时", Replacement:="0.25天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="3小时", Replacement:="0.375天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="4小时", Replacement:="0.5天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="5小时", Replacement:="0.625天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="6小时", Replacement:="0.75天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="7小时", Replacement:="0.875天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Function

Function replaceDay()
    Cells.Replace What:="休1", Replacement:="倒休0.125天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="休2", Replacement:="倒休0.25天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="休3", Replacement:="倒休0.375天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="休4", Replacement:="倒休0.5天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="休5", Replacement:="倒休0.625天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="休6", Replacement:="倒休0.75天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="休7", Replacement:="倒休0.875天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事1", Replacement:="事假0.125天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事2", Replacement:="事假0.25天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事3", Replacement:="事假0.375天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事4", Replacement:="事假0.5天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事5", Replacement:="事假0.625天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事6", Replacement:="事假0.75天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="事7", Replacement:="事假0.875天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="退1", Replacement:="早退0.125天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="早早退0.125天天", Replacement:="早退1天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷1", Replacement:="迟到0.125天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷7", Replacement:="旷工0.875天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工0.125天", Replacement:="迟到0.125天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷4", Replacement:="迟到0.5天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="旷工0.5天", Replacement:="迟到0.5天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病1", Replacement:="病假0.125天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病2", Replacement:="病假0.25天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病3", Replacement:="病假0.375天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病4", Replacement:="病假0.5天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病5", Replacement:="病假0.625天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病6", Replacement:="病假0.75天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="病7", Replacement:="病假0.875天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值1", Replacement:="值班0.125天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值2", Replacement:="值班0.25天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值3", Replacement:="值班0.375天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值4", Replacement:="值班0.5天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值5", Replacement:="值班0.625天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值6", Replacement:="值班0.75天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="值7", Replacement:="值班0.875天", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Function
Function oneDayModify()
    Dim monthDay, monthDayModify As Variant
    Dim tmpRow As Integer
    Dim tmpName As String
    monthDay = Array("L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ")
    monthDayModify = Array("H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ")
'    For k = 0 To UBound(monthDay)
'        Debug.Print monthDay(k)
'    Next
    For I = 0 To 30 'day
        For j = 0 To 176 'person
            If InStr(Sheets("修正").Range(monthDayModify(I) & j + 6).Value, "补") > 0 Then
                tmpName = Sheets("修正").Range("E" & j + 6).Value
                Debug.Print tmpName
                Sheets("2月修正").Select
                
                Cells.Find(What:=tmpName, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                    , MatchByte:=False, SearchFormat:=False).Activate
                tmpRow = Selection.Row
                
                Range(monthDay(I) & tmpRow).Select
                Sheets("2月修正").Range(monthDay(I) & tmpRow).Value = Sheets("修正").Range(monthDayModify(I) & j + 6).Value
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    If .Color <> 65535 Then
                        .Color = 65535
                    End If
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'callDebug
                tmpName = ""
                tmpRow = 0
            End If
        Next
    Next
End Function


Function testProjectDay()
    For I = 1 To 1
        projectDay (I)
    Next
End Function

Function projectDay(day As Integer)
    Dim monthDay  As Variant
    Dim formular, tmpName, timeStr As String
    Dim isProject, needNext As Boolean
    monthDay = Array("D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH")
    isProject = False
    needNext = True
    If day < 10 Then
        timeStr = "25-01-0" & day
    Else
        timeStr = "25-01-" & day
    End If
    For I = 1 To 20 '暂定人数
        tmpName = Sheets("H").Range("B" & I + 199).Value
        On Error GoTo ErrorHandler
        If tmpName = "" Then
            Exit For
        Else
            Sheets("add").Select
            Columns("G:G").Select
            Selection.Replace What:="星期*", Replacement:="--", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            ActiveSheet.Range("$A:$T").AutoFilter Field:=7, Criteria1:=timeStr & " --" '    $A$3:$T$5654
            Cells.Find(What:=tmpName, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                , MatchByte:=False, SearchFormat:=False).Activate
            tmpRow = Selection.Row
            If InStr(Range("K" & tmpRow).Value, "恒有源科技发展有限公司") > 0 Then
                 While needNext
                    tmpRow = tmpRow + 1
                    If Range("A" & tmpRow).Value = tmpName And Left(Range("G" & tmpRow).Value, 8) = timeStr Then
                        needNext = True
                    Else
                        needNext = False
                        tmpRow = tmpRow - 1
                    End If
                Wend
                If Not InStr(Range("K" & tmpRow).Value, "恒有源科技发展有限公司") > 0 Then
                    isProject = True
                End If
            Else
                isProject = True
            End If
        End If
        Sheets("H").Select
        If isProject Then
            Range(monthDay(day - 1) & I + 199).Value = "工程"
        Else
            Range(monthDay(day - 1) & I + 199).Value = "办公"
        End If
        isProject = False
        needNext = True
ErrorHandler:
        Sheets("H").Select
        On Error GoTo -1
    Next
End Function

Function workDays(day As Integer)
'Function workDays()
    Dim monthDay  As Variant
    Dim formular As String
    monthDay = Array("F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ")
    formular = "=COUNTIF($F308:$" & monthDay(day - 1) & "308,$C$307)"
    Sheets("H").Select
    Range("B308").Select
    ActiveCell.Formula = formular
    Range("B308").Select
    Selection.Autofill Destination:=Range("B308:B480")
    Range("B308").Select
End Function

Function saveNewFile(day As Integer, weekDay As String)
'Function saveNewFile()
    'Dim day As Integer
    'day = 6
    Dim newFileName, shortName, pdfFileName As String
    newFileName = "D:\调度\团队号\钉钉\考勤1月" & day & "号.xlsx"
    shortName = "考勤1月" & day & "号.xlsx"
    pdfFileName = "D:\调度\团队号\钉钉\考勤1月" & day & "号.pdf"
    Workbooks.Add
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=newFileName, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Windows("1月考勤.xlsm").Activate
    If InStr(weekDay, "日") > 0 Or InStr(weekDay, "六") > 0 Then
        Sheets(Array("d7", "mo")).Select
        Sheets(Array("d7", "mo")).Copy Before:=Workbooks(shortName).Sheets(1)
        Windows(shortName).Activate
        Sheets("sheet1").Delete
        Application.DisplayAlerts = True
        Call clearSheets
        ActiveWorkbook.Save
        Sheets(Array("d7", "mo")).Select
    '    ChDir "D:\调度\团队号\钉钉"
    '    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    '        pdfFileName, Quality:=xlQualityStandard, IncludeDocProperties _
    '        :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("d7").Select
    Else
        Sheets(Array("d", "mo")).Select
        Sheets(Array("d", "mo")).Copy Before:=Workbooks(shortName).Sheets(1)
        Windows(shortName).Activate
        Sheets("sheet1").Delete
        Application.DisplayAlerts = True
        Call clearSheets
        ActiveWorkbook.Save
        Sheets(Array("d", "mo")).Select
    '    ChDir "D:\调度\团队号\钉钉"
    '    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    '        pdfFileName, Quality:=xlQualityStandard, IncludeDocProperties _
    '        :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("d").Select
    End If
    ActiveWorkbook.Close
    Windows("1月考勤.xlsm").Activate
    Sheets("1").Select
    Sheets("1月").Select
End Function
Function testOneDay()
    Call oneDay(1, "周六")

End Function

Function oneDay(day As Integer, weekDay As String)
    Dim oneDayColumnStart, firstCellStart, secondCellStart, dayNumber As Integer
    Dim firstCell, secondCell, tmpStr As String
    oneDayColumnStart = GetColumnNum("K") '月度表1号列
    firstCellStart = GetColumnNum("I") '计算表的数据区，不变
    secondCellStart = firstCellStart + 1
    dayNumber = day
    Call workDays(dayNumber)
    
    Sheets("1月").Select
    totalLine = 176
    LogNum = 0
    For Line = 7 To totalLine
        Delayms (0.1)
        'If Line = 9 Then callDebug
        For col = dayNumber To dayNumber '
            tmpStr = ""
            Range(GetColumnName(oneDayColumnStart + (col - 1)) & Line).Select
            firstCell = Trim(Sheets("计算").Range(GetColumnName(firstCellStart + (col - 1) * 2) & Line).Value)
            secondCell = Trim(Sheets("计算").Range(GetColumnName(secondCellStart + (col - 1) * 2) & Line).Value)
            cellName = Selection.Address
            ColumnName = GetColumnName(Selection.Column)
            Debug.Print Line & "：" & Range("D" & Selection.Row).Value & "--行号--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
            
            Debug.Print firstCell & "-" & secondCell
            'If (firstCell = "休" And secondCell = "休") Or (firstCell = 0 And secondCell = 0) Then
            If (firstCell = secondCell) Then
                If (firstCell = "休" And secondCell = "休") Then
                    tmpStr = ""
                ElseIf firstCell = "正常" And secondCell = "正常" Then
                    tmpStr = "√"
                Else
                    tmpStr = firstCell
                End If
            Else
                If secondCell = "休" Then
                    tmpStr = firstCell
                ElseIf firstCell = "休" Then
                    tmpStr = secondCell
                ElseIf IsNumeric(firstCell) And IsNumeric(secondCell) Then
                    tmpStr = firstCell - secondCell
                ElseIf IsNumeric(firstCell) And secondCell = "正常" Then
                    tmpStr = firstCell
                ElseIf IsNumeric(firstCell) And secondCell <> "正常" Then
                    tmpStr = firstCell & Chr("10") & secondCell
                ElseIf IsNumeric(secondCell) And firstCell = "正常" Then
                    tmpStr = secondCell
                ElseIf IsNumeric(secondCell) And firstCell <> "正常" Then
                    tmpStr = firstCell & Chr("10") & secondCell
                ElseIf firstCell = "正常" And secondCell <> "正常" Then
                    If secondCell = "倒休" Then
                        tmpStr = "休4"
                    ElseIf secondCell = "事假" Then
                        tmpStr = "事4"
                    ElseIf secondCell = "病假" Then
                        tmpStr = "病4"
                    ElseIf secondCell = "年假" Then
                        tmpStr = "年4"
                    ElseIf firstCell = "未签到" Then
                        tmpStr = "未签到"
                    ElseIf secondCell = "旷1" Then
                        tmpStr = "退1"
                    ElseIf secondCell = "旷4" Then
                        tmpStr = "退4"
                    Else
                        tmpStr = secondCell
                    End If
                ElseIf firstCell <> "正常" And secondCell = "正常" Then
                    If firstCell = "倒休" Then
                        tmpStr = "休4"
                    ElseIf secondCell = "未签退" Then
                        tmpStr = "未签退"
                    ElseIf firstCell = "事假" Then
                        tmpStr = "事4"
                    ElseIf firstCell = "病假" Then
                        tmpStr = "病4"
                    ElseIf firstCell = "年假" Then
                        tmpStr = "年4"
                    Else
                        tmpStr = firstCell
                    End If
                ElseIf firstCell = "" And secondCell = "倒休" Then
                    tmpStr = "休4"
                ElseIf firstCell = "倒休" And secondCell = "" Then
                    tmpStr = "休4"
                ElseIf firstCell = "值班" And secondCell <> "值班" Then
                    If secondCell = "" Or secondCell = 0 Then
                        tmpStr = "值4"
                    Else
                        tmpStr = secondCell
                    End If
                ElseIf firstCell <> "值班" And secondCell = "值班" Then
                    If firstCell = "" Or firstCell = 0 Then
                        tmpStr = "值4"
                    Else
                        tmpStr = firstCell
                    End If
                ElseIf firstCell <> "值班请假" And secondCell = "值班请假" Then
                    tmpStr = firstCell
                ElseIf firstCell <> "值班旷工" And secondCell = "值班旷工" Then
                    tmpStr = firstCell
                Else
                    tmpStr = firstCell & Chr("10") & secondCell
                End If
            End If

            Delayms (0.01)
           ' If Line = 7 Then Range(GetColumnName(oneDayColumnStart + col - 1) & 6) = "1月" & col & "日"
           Range(GetColumnName(oneDayColumnStart + (col - 1)) & Line).Select
           Selection.Value = tmpStr
           If InStr(tmpStr, Chr("10")) > 0 Then callDebug '================
           'Range(GetColumnName(oneDayColumnStart + (col - 1)) & Line).Select
           If IsNumeric(tmpStr) > 0 Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 14351096
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
           ElseIf IsNumeric(tmpStr) < 0 Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "病1") Or InStr(tmpStr, "病2") Or InStr(tmpStr, "病3") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16751103
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "病假") Then
                With Selection.Font
                    .Color = -39169
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "×") Then
                With Selection.Font
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "出差") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 10092543
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With Selection.Font
                    .Color = -16750951
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "丧假") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark2
                    .TintAndShade = -9.99786370433668E-02
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "值4") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 15773696
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "事1") Or InStr(tmpStr, "事2") Or InStr(tmpStr, "事3") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16764108
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "事假") Then
                With Selection.Font
                    .Color = -52327
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "休1") Or InStr(tmpStr, "休2") Or InStr(tmpStr, "休3") Or InStr(tmpStr, "休4") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 5296274
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "倒休") Then
                With Selection.Font
                    .Color = -11489280
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "年") Then
                With Selection.Font
                    .Color = -65536
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "补") Then
                With Selection.Font
                    .Color = -11489280
                    .TintAndShade = 0
                End With
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "未") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With Selection.Font
                    .Color = -16776961
                    .TintAndShade = 0
                End With
            End If
        Next
    Next Line
    Call replaceDay
    Call projectDay(dayNumber)
'    Call formatResult
'    Call formatResult2
    Call saveNewFile(dayNumber, (weekDay))
    ActiveWorkbook.Save
End Function


Public Function Delayms(lngTime As Long)
    Dim StartTime As Single
    Dim CostTime As Single
    StartTime = Timer
     
    Do While (Timer - StartTime) * 1000 < lngTime
        DoEvents
    Loop
End Function


'1 ?通过列名称转换成对应的列号?
Function GetColumnNum(ByVal ColumnN As String) As Integer
    Dim result As Integer, First As Integer, Last As Integer
    result = 1
     If Trim(ColumnN) <> "" Then
         If Len(ColumnN) = 1 Then
           result = Asc(UCase(ColumnN)) - 64
         ElseIf Len(ColumnN) = 2 Then
           If UCase(ColumnN) > "IV" Then ColumnN = "IV"
           First = Asc(UCase(Left(ColumnN, 1))) - 64
            Last = Asc(UCase(Right(ColumnN, 1))) - 64
            result = First * 26 + Last
        End If
     End If
     GetColumnNum = result
End Function
'测试:

Function TestGetColumnNum()
    Dim ColumnNum As Integer
     ColumnNum = GetColumnNum("ET")
     Debug.Print ColumnNum
     'MsgBox ColumnNum, vbInformation, "测试"
 End Function
 
'2 ?通过列号转换成对应的列名称?
Function GetColumnName(ByVal ColumnNu As Integer) As String
     Dim First As Integer, Last As Integer
     Dim result As String
     If ColumnNu < 27 Then
        result = Chr(ColumnNu + 64)
     Else
        If ColumnNu > 256 Then ColumnNu = 256
         First = Int(ColumnNu / 26)
         Last = ColumnNu - (First * 26)
         If First * 26 = ColumnNu Then First = First - 1
        If First > 0 Then
            result = Chr(First + 64)
        End If
         If Last > 0 Then
            result = result & Chr(Last + 64)
        ElseIf Last = 0 Then
            result = result & "Z"
        End If
    End If
    GetColumnName = result
 End Function
'测试:
'Function TestGetColumnName()
'     Dim ColumnName As String
'     ColumnName = GetColumnName(54)
'    MsgBox ColumnName, vbInformation, "测试"
' End Function
'3 ?说明
'在以上两个函数中，如果输入的参数大于Excel的最大列号"IV"(256),则返回的值为最大的列数。
'————————————————
'版权声明：本文为CSDN博主「xuanxingmin」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
'原文链接：https://blog.csdn.net/xuanxingmin/article/details/2582861

Function statusTest()
    Sheets("day").Select
    Dim coName(), StatusName() As Variant, temp, rangePosi As String, flagRow, flagCol As Integer
    
    coName() = Range("A6:A23").Value
    StatusName() = Range("I5:BB5").Value
    
    temp = "财务部"
    
    On Error Resume Next
    flagRow = WorksheetFunction.Match(temp, coName, 0) + 5 '上面有5行
    Debug.Print flagRow
        
    temp = "加班"
    On Error Resume Next
    flagCol = WorksheetFunction.Match(temp, StatusName, 0) + 8 '前面有8列
    rangePosi = GetColumnName(flagCol) & flagRow
    Debug.Print rangePosi
    
    Debug.Print Range(rangePosi).Value
    
End Function

Function IsWeekNormal(ByVal searchValue As String) As Boolean
    Sheets("计算").Select
    For I = 460 To 480
        If Range("D" & I).Value = "" Then
            IsWeekNormal = False
            Exit Function
        End If
        If searchValue = Range("D" & I).Value Then
            IsWeekNormal = True
            Exit Function
        End If
    Next I
    
    IsWeekNormal = False
End Function

Function testRemove()
    Debug.Print removeNonDigits1("g一1.207")
End Function

Function KeepNumbersAndDecimals(strInput As String) As String
    ' 使用正则表达式替换非数字和非小数点的字符为空
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .Pattern = "[^\d.]"
        KeepNumbersAndDecimals = .Replace(strInput, vbNullString)
    End With
End Function

Function callDebug()
    Application.EnableCancelKey = xlInterrupt '设置取消键为中断状态
    Debug.Assert False '将此行作为断点放置在需要调试的位置上
End Function

Function SortStr(str As String) As String
    Dim arr As Variant

    '定义要排序的数组
'    arr = Array(5, 2, 8, 1, 9)
    arr = Split(str, Chr(10))

    '调用 Sort 函数对数组进行排序（默认为升序）
    Call Sort(arr)

    '输出排序后的结果
    For I = LBound(arr) To UBound(arr)
        Debug.Print arr(I)
        If I = LBound(arr) Then
            SortStr = arr(I)
        Else
            SortStr = SortStr & Chr(10) & arr(I)
        End If
    Next I
End Function
 
'自定义的 Sort 过程
Function Sort(ByRef arr As Variant)
    Dim tempArr As Variant
    Dim I As Integer, j As Integer
    
    '将传入的数组赋值给临时变量
    tempArr = arr
    
    '通过比较相邻元素并交换位置来完成排序
    For I = LBound(tempArr) To UBound(tempArr) - 1
        For j = I + 1 To UBound(tempArr)
            If tempArr(j) < tempArr(I) Then
                SwapElements tempArr, I, j
            End If
        Next j
    Next I
    
    '将排序好的数组重新赋值给原始数组
    arr = tempArr
End Function
 
'交换两个元素的位置
Sub SwapElements(ByRef arr As Variant, ByVal index1 As Integer, ByVal index2 As Integer)
    Dim temp As Variant
    
    temp = arr(index1)
    arr(index1) = arr(index2)
    arr(index2) = temp
End Sub

Sub ChangeStrTypeNumber()
    n = 0
    tmp = c.Value
    If IsNumeric(tmp) Then
        c.NumberFormatLocal = ""
        c.Value = Val(tmp)
        n = n + 1
    ElseIf Right(tmp, 1) = "%" And IsNumeric(Left(tmp, Len(tmp) - 1)) Then
        If InStr(tmp, ".") = 0 Then
            c.NumberFormatLocal = "0%"
        Else
            pos = Len(tmp) - InStr(tmp, ".") - 1
            c.NumberFormatLocal = "0." & WorksheetFunction.Rept("0", pos) & "%"
        End If
        c.Value = Val(Left(tmp, Len(tmp) - 1)) / 100
        n = n + 1
    End If
    If n > 0 Then MsgBox "文本型数字转换为数值成功！", vbOKOnly, "成功" Else MsgBox "未检测到文本型数字！", vbOKOnly + vbCritical, "错误"

End Sub


Sub ChangeStrTypeNumber1()
    Dim rng As Range
    Set rng = Intersect(ActiveSheet.UsedRange, Selection)
    If rng Is Nothing Then Exit Sub
    n = 0
    For Each c In rng
        If VarType(c) = 8 And c.Value <> "" Then 'VarTye:7-Date,8-String
            tmp = c.Value
            If IsNumeric(tmp) Then
                c.NumberFormatLocal = ""
                c.Value = Val(tmp)
                n = n + 1
            ElseIf Right(tmp, 1) = "%" And IsNumeric(Left(tmp, Len(tmp) - 1)) Then
                If InStr(tmp, ".") = 0 Then
                    c.NumberFormatLocal = "0%"
                Else
                    pos = Len(tmp) - InStr(tmp, ".") - 1
                    c.NumberFormatLocal = "0." & WorksheetFunction.Rept("0", pos) & "%"
                End If
                c.Value = Val(Left(tmp, Len(tmp) - 1)) / 100
                n = n + 1
            End If
        End If
    Next c
    If n > 0 Then MsgBox "文本型数字转换为数值成功！", vbOKOnly, "成功" Else MsgBox "未检测到文本型数字！", vbOKOnly + vbCritical, "错误"
    Set rng = Nothing
End Sub
