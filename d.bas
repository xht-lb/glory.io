Dim cellName, vStr, ColumnName As String
Dim recordLine, LogNum As Integer
Dim cd, zt, kg, cc, dx, zcx, sj, bj, qk, rz, lz, nx, bc, bt, bk, wqd, wqt, bqd, bqt  As Integer '�ٵ������ˡ�������������ݡ������ݡ��¼١����١�ȱ������ְ����ְ�����ݡ����١����ˡ���������ǩ������ǩ��
Dim cds, zts, kgs, ccs, dxs, zcxs, sjs, bjs, qks, rzs, lzs, nxs, bcs, bts, bks, wqds, wqts, bqds, bqts As Integer '�ٵ������ˡ�������������ݡ������ݡ��¼١����١�ȱ������ְ����ְ������
Dim cdStr, ztStr, kgStr, ccStr, dxStr, zcxStr, sjStr, bjStr, qkStr, rzStr, lzStr, nxStr, bcStr, btStr, bkStr, wqdStr, wqtStr, bqdStr, bqtStr As String '�ٵ������ˡ�������������ݡ������ݡ��¼١����١�ȱ������ְ����ְ������
'initialSheets monthDataClear �滻*�� monthData projectDay dingtalk(totalLine) oneDay(totalLine) oneDay(N) workDays(line)
'176 170 ������ͷ
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
        
        weekDay = dayReportBook((I)) '�������ݵ�ͬʱ����������
        Sheets("day").Select
        '���Ű����ʵΪ�������ڣ���������ȷλ��
        Range("I6:I23").Select
        Range("I23").Activate
        Selection.Copy
        Range("F6").Select
        ActiveSheet.Paste
        Range("I31").Select '���Ű���ڹ�ʽ���Ƶ���Ӧ��Ԫ��
        Application.CutCopyMode = False
        Selection.Copy
        Range("I6:I23").Select
        ActiveSheet.Paste
        
        '���Ƶ���Ӧ���������ܱ�
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
        Sheets("M").Range("B1").Value = "�й�����Դ���ſ�����ͳ�Ʊ�����������2025��1��" & I & "��"
        Sheets("M").Range("B27").Value = "�й�����Դ���ſ�����ͳ�Ʊ�������������2025��1��" & I & "��"
    
        Sheets(sheetName).Range("B1").Value = "�й�����Դ���ſ�����ͳ�Ʊ�����������2025��1��" & I & "��"
        Sheets(sheetName).Range("B27").Value = "�й�����Դ���ſ�����ͳ�Ʊ�������������2025��1��" & I & "��"
        Call replaceDayHour

            

        Call lessName((sheetName))
        Call dealNewD((I), (weekDay))
        Call projectDay((I))
        Call oneDay((I), (weekDay))
        
        Call doublePerLine
        Call dayRowHigh
        Sheets("day").Select
        ChDir "D:\����\�ŶӺ�\����"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "D:\����\�ŶӺ�\����\1��" & I & "�տ���.pdf", Quality:=xlQualityStandard, IncludeDocProperties _
            :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        'ActiveWorkbook.Save
    Next
End Function

Function dingtalk(day As Integer)

'1 A  ����   '2 B  ������  '3 C  ����  '4 D  ����  '5  E ְλ   '6 F  UserID  '7 G  ����  '8 H  workDate  '9 I  ���
'10 J �ϰ�1��ʱ��  '11 K �ϰ�1�򿨽��  '12 L �°�1��ʱ��  '13 M �°�1�򿨽��
'14 N �ϰ�2��ʱ��  '15 O �ϰ�2�򿨽��  '16 P �°�2��ʱ��  '17 Q �°�2�򿨽��
'18 R �ϰ�3��ʱ��  '19 S �ϰ�3�򿨽��  '20 T �°�3��ʱ��  '21 U �°�3�򿨽��
'22 V ������������   '23 W ��������       '24 X ��Ϣ����       '25 Y ����ʱ��
'26 Z �ٵ�����       '27 AA �ٵ�������     '28 AB ���سٵ�����   '29 AC ���سٵ�������  '30 AD �����ٵ�����
'31 AE ���˴���       '32 AF ���˷�����     '33 AG �ϰ�ȱ������   '34 AH �°�ȱ������    '35 AI ��������
'36 AJ ��������       '37 AK ���ʱ��       '38 AL �Ӱ���ʱ��     '39 AM �Ӱ�ʱ����ת���ݣ�  �����գ�ת���ݣ�
'40 AN     ��Ϣ�� (ת����) '41 AO     �ڼ��� (ת����) '42 AP �Ӱ�ʱ����ת�Ӱ�ѣ�    �����գ�ת�Ӱ�ѣ�
'43 AQ     ��Ϣ�� (ת�Ӱ��) '44  AR    �ڼ��� (ת�Ӱ��)

    Sheets("����").Select
    totalLine = 170
    Dim posiPre, posi, tmpName, sheetName, tmpColumn, shiftName, amStatus, pmStatus, shiftGroup, weekDay, thisAMStatus, thisPMStatus, relavition, coName As String
    posiPre = "A"
    posi = GetColumnName(GetColumnNum(posiPre) + 10) '�ϰ�1�򿨽�� K
    Debug.Print posi
    Dim lineNum, endLine, tmpRow, lateMins, lateMore, lateLeave, ampmStatusRow, thisDay As Integer '�ٵ������سٵ������˷�����
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
            If Trim(Sheets("����").Range(tmpColumn & I).Value) = "" Then
                thisPMStatus = "": thisAMStatus = thisPMStatus
                
                'If tmpName <> "Ѧ����" And tmpName <> "ɳ����" Then
                If tmpName <> "Ѧ����" And tmpName <> "ɳ����" And tmpName <> "�Ծ�" And tmpName <> "������" Then
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
                    If InStr(amStatus, "��Ϊ����") > 0 Then amStatus = "����"
                    If InStr(pmStatus, "��Ϊ����") > 0 Then pmStatus = "����"
                    
                    lateMins = --Range("AA" & tmpRow).Value
                    lateMore = --Range("AC" & tmpRow).Value
                    lateLeave = --Range("AF" & tmpRow).Value '���˷�����
                    relavition = Range("V" & tmpRow).Value
                    coName = Range("C" & tmpRow).Value
                    Debug.Print tmpName & ">����-��ǰ��-λ����>" & I - 6 & "-" & tmpRow & ">���>" & shiftName & ">�ϰ�򿨽��>" & amStatus & ">�°�򿨽��>" & pmStatus
                    
                    Sheets("����").Select
    
                    If tmpName = "������" Then
                        If amStatus = "" And pmStatus = "" Then
                            thisPMStatus = "��": thisAMStatus = thisPMStatus
                        ElseIf amStatus = "����" And pmStatus = "����" Then
                            thisPMStatus = "����": thisAMStatus = thisPMStatus
                        ElseIf amStatus = "����" And pmStatus = "δǩ��" Then
                            thisPMStatus = "δǩ��": thisAMStatus = "����"
                        ElseIf amStatus = "δǩ��" And pmStatus = "����" Then
                            thisPMStatus = "����": thisAMStatus = "δǩ��"
                        ElseIf amStatus = "����" And pmStatus = "" Then
                            thisAMStatus = "����": thisPMStatus = "δǩ��"
                        Else
                            callDebug
                        End If
                    ElseIf amStatus = "���" And pmStatus = "���" Then
                        thisAMStatus = Left(relavition, 2): thisPMStatus = Left(relavition, 2)
                    ElseIf amStatus = "���" And pmStatus = "δǩ��" Then
                        thisAMStatus = "����": thisPMStatus = pmStatus
                    ElseIf amStatus = "δǩ��" And pmStatus = "���" Then
                        thisAMStatus = amStatus: thisPMStatus = "����"
                    ElseIf amStatus = "ȱ��" And pmStatus = "ȱ��" And relavition <> "" Then
                        thisAMStatus = Left(relavition, 2): thisPMStatus = Left(relavition, 2)
                        'callDebug
                    ElseIf (amStatus = pmStatus = "����") Or (amStatus = "����" And pmStatus = "����") Or (pmStatus = "����" And amStatus = "����") Or (pmStatus = "����" And amStatus = "����") Then
                        If Left(relavition, 2) = "����" Then
                            thisAMStatus = "����": thisPMStatus = "����"
                        ElseIf Left(relavition, 2) = "���" Then
                            thisPMStatus = "����": thisAMStatus = thisPMStatus
                        Else
                            callDebug
                            thisPMStatus = "����": thisAMStatus = thisPMStatus
                        End If
                    ElseIf amStatus = "���سٵ�" Then
                        If Sheets(sheetName).Range("AC" & tmpRow).Value < 60 Then
                            thisAMStatus = "��1"
                        ElseIf Sheets(sheetName).Range("AC" & tmpRow).Value < 120 Then
                            thisAMStatus = "��4"
                        Else
    '                        ActiveWorkbook.Save
                            Debug.Print tmpName & ">����-��ǰ��-λ����>" & I & "-" & tmpRow & ">���>" & shiftName & ">�ϰ�򿨽��>" & amStatus & ">�°�򿨽��>" & pmStatus
                            callDebug
                        End If
                        
                        If pmStatus = "ȱ��" Then
                            thisPMStatus = "δǩ��"
                        ElseIf pmStatus = "����" Then
                            thisPMStatus = "����"
                        ElseIf pmStatus = "����" Then
                            If Sheets(sheetName).Range("AF" & tmpRow).Value < 30 Then '���˷�����
                                thisPMStatus = "����"
                            ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 60 Then '���˷�����
                                thisPMStatus = "��1"
                            ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 120 Then '���˷�����
                                thisPMStatus = "��4"
                            Else
                            
                                thisPMStatus = "����1��"
                                thisAMStatus = thisPMStatus
                                'ActiveWorkbook.Save
                                Debug.Print tmpName & ">����-��ǰ��-λ����>" & I & "-" & tmpRow & ">���>" & shiftName & ">�ϰ�򿨽��>" & amStatus & ">�°�򿨽��>" & pmStatus
                                callDebug
                            End If
                        ElseIf pmStatus = "���" Then
                            callDebug
                            thisPMStatus = Left(relavition, 2)
                        ElseIf pmStatus = "����" Then
                            If Left(relavition, 2) = "����" Then
                                 thisPMStatus = "����"
                            ElseIf Left(relavition, 2) = "���" Then
                                 thisPMStatus = "����"
                            Else
                                callDebug
                            End If
                        Else
                            callDebug
                        End If
                    ElseIf amStatus = "�����ٵ�" Then
                        thisPMStatus = "�ٵ�1��": thisAMStatus = thisPMStatus
                    ElseIf amStatus = "���" And pmStatus = "δ��" Then
                        thisPMStatus = Left(relavition, 2): thisAMStatus = thisPMStatus
                    ElseIf amStatus = "�ٵ�" And pmStatus = "����" Then
                        thisAMStatus = "�ٵ�"
                        thisPMStatus = Left(relavition, 2)
                        If thisPMStatus = "���" Then thisPMStatus = "����"
                    ElseIf amStatus = "�ٵ�" And pmStatus = "ȱ��" Then
                        thisAMStatus = "�ٵ�"
                        thisPMStatus = "δǩ��"
                    ElseIf amStatus = "�ٵ�" Then
                        If pmStatus = "����" Then
                            thisAMStatus = "�ٵ�"
                            thisPMStatus = "����"
                        ElseIf pmStatus = "δ��" Then
                            thisAMStatus = "�ٵ�"
                            thisPMStatus = "δǩ��"
                        ElseIf pmStatus = "����" Then
                            thisAMStatus = "�ٵ�"
                            If Sheets(sheetName).Range("AF" & tmpRow).Value < 30 Then '���˷�����
                                thisPMStatus = "����"
                            ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 60 Then '���˷�����
                                thisPMStatus = "��1"
                            ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 120 Then '���˷�����
                                thisPMStatus = "��4"
                            Else
                            
                                thisPMStatus = "����1��"
                                thisAMStatus = thisPMStatus
                                'ActiveWorkbook.Save
                                Debug.Print tmpName & ">����-��ǰ��-λ����>" & I & "-" & tmpRow & ">���>" & shiftName & ">�ϰ�򿨽��>" & amStatus & ">�°�򿨽��>" & pmStatus
                                callDebug
                            End If
                        Else
                            callDebug
                        End If
                    ElseIf amStatus = "ȱ��" And pmStatus = "����" Then
                        thisAMStatus = "δǩ��"
                        thisPMStatus = Left(relavition, 2)
                        If thisPMStatus = "���" Then thisPMStatus = "����"
                    ElseIf pmStatus = "ȱ��" And amStatus = "����" Then
                        thisPMStatus = "δǩ��"
                        thisAMStatus = Left(relavition, 2)
                        If thisPMStatus = "���" Then thisPMStatus = "����"
                    ElseIf pmStatus = "����" Then
                        If Sheets(sheetName).Range("AF" & tmpRow).Value < 30 Then '���˷�����
                            thisPMStatus = "����"
                        ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 60 Then '���˷�����
                            thisPMStatus = "��1"
                        ElseIf Sheets(sheetName).Range("AF" & tmpRow).Value < 120 Then '���˷�����
                            thisPMStatus = "��4"
                        Else
                            thisPMStatus = "����1��"
                            thisAMStatus = thisPMStatus
    '                        ActiveWorkbook.Save
                            Debug.Print tmpName & ">����-��ǰ��-λ����>" & I & "-" & tmpRow & ">���>" & shiftName & ">�ϰ�򿨽��>" & amStatus & ">�°�򿨽��>" & pmStatus
                            
                            callDebug
                        End If
                        
                        If amStatus = "����" Then
                            thisAMStatus = "����"
                        Else
                            callDebug
                        End If
                        
                        If thisPMStatus = "����1��" Then
                            thisAMStatus = thisPMStatus
                        Else
                            callDebug
                        End If
                        
                    ElseIf 1 Then
                        Dim ampmStatus() As Variant
                        Dim flagRow As Integer
                        ampmStatusRow = 510
                        ampmStatus() = Range("D481:D" & ampmStatusRow).Value
                        flagRow = WorksheetFunction.Match(amStatus & "-" & pmStatus, ampmStatus, 0) + 480 '������5��
                        'Debug.Print flagRow
                        thisAMStatus = Range("E" & flagRow).Value
                        thisPMStatus = Range("F" & flagRow).Value
    
                    Else
                        Debug.Print amStatus & pmStatus
                        callDebug
                    End If
        
                    If weekDay = "����" Or weekDay = "����" Or isFday Then
                        If tmpName = "ɳ����" Then
                        'If tmpName = "³����" Or tmpName = "Ѧ����" Then
                            thisStatus = "��" '1.9-2.6
                        ElseIf tmpName = "������" Or tmpName = "������" Or tmpName = "�Ծ�" Or tmpName = "������" Then
                            thisStatus = "��ְ"
                        End If
                    
                        If thisAMStatus = "����" Then
    '                        If tmpName = "�´���" Then
    '                            thisAMStatus = "����"
    '                        ElseIf tmpName = "��ȫϲ" Then
    '                            thisAMStatus = "ֵ��"
                            If IsWeekNormal(tmpName) Then
                                thisAMStatus = "����"
                            ElseIf InStr(shiftName, "ֵ") > 0 Or Not InStr(shiftName, "��") Then
                                thisAMStatus = "ֵ��"
                            ElseIf shiftGroup = "�¶�ֵ��" And Left(relavition, 2) = "�Ӱ�" Then
                               If shiftName = "��Ϣ" Then
                                    thisAMStatus = "�Ӱ�"
                                Else
                                    thisAMStatus = "ֵ��"
                                End If
                            ElseIf shiftGroup = "�칫��" And Left(relavition, 2) = "�Ӱ�" Then
                                If InStr(coName, "ԴȪ") > 0 Or InStr(coName, "��װ") > 0 Or InStr(coName, "һ��") > 0 Then
                                    thisAMStatus = "�Ӱ�"
                                Else
                                    thisAMStatus = "ֵ��"
                                End If
                            ElseIf shiftGroup = "����" And Left(relavition, 2) = "�Ӱ�" Then
                                If shiftName = "��Ϣ" Then
                                    thisAMStatus = "�Ӱ�"
                                Else
                                    thisAMStatus = "ֵ��"
                                End If
                            ElseIf InStr(shiftGroup, "��ά") > 0 And Left(relavition, 2) = "�Ӱ�" Then
                                If shiftName = "��Ϣ" Then
                                    thisAMStatus = "�Ӱ�"
                                Else
                                    thisAMStatus = "ֵ��"
                                End If
                            Else
                                callDebug
                                Debug.Print tmpName & ">����-��ǰ��-λ����>" & I & "-" & tmpRow & ">���>" & shiftName & ">�ϰ�򿨽��>" & amStatus & ">�°�򿨽��>" & pmStatus
                            End If
                        End If
                        
                        If thisPMStatus = "����" Then
    '                        If tmpName = "�´���" Then
    '                            thisPMStatus = "����"
    '                        ElseIf tmpName = "��ȫϲ" Then
    '                            thisPMStatus = "ֵ��"
                            If IsWeekNormal(tmpName) Then
                                thisPMStatus = "����"
                            ElseIf InStr(shiftName, "ֵ") > 0 Or Not InStr(shiftName, "��") Then
                                thisPMStatus = "ֵ��"
                            ElseIf shiftGroup = "�¶�ֵ��" And Left(relavition, 2) = "�Ӱ�" Then
                               If shiftName = "��Ϣ" Then
                                    thisPMStatus = "�Ӱ�"
                                Else
                                    thisPMStatus = "ֵ��"
                                End If
                            ElseIf shiftGroup = "�칫��" And Left(relavition, 2) = "�Ӱ�" Then
                                If InStr(coName, "ԴȪ") > 0 Or InStr(coName, "��װ") > 0 Or InStr(coName, "һ��") > 0 Then
                                    thisPMStatus = "�Ӱ�"
                                Else
                                    thisPMStatus = "ֵ��"
                                End If
                            ElseIf shiftGroup = "����" And Left(relavition, 2) = "�Ӱ�" Then
                                If shiftName = "��Ϣ" Then
                                    thisPMStatus = "�Ӱ�"
                                Else
                                    thisPMStatus = "ֵ��"
                                End If
                            ElseIf InStr(shiftGroup, "��ά") > 0 And Left(relavition, 2) = "�Ӱ�" Then
                                If shiftName = "��Ϣ" Then
                                    thisPMStatus = "�Ӱ�"
                                Else
                                    thisPMStatus = "ֵ��"
                                End If
                            Else
                                callDebug
                            End If
                        End If
                        If (Not IsWeekNormal(tmpName)) Then
                            If thisAMStatus = "����" And thisPMStatus = "����" Then
                                     thisPMStatus = "ֵ�����": thisAMStatus = thisPMStatus
                            ElseIf (thisAMStatus = "����" And thisPMStatus = "����") Or (thisAMStatus = "����" And thisPMStatus = "����") Then
                                     thisPMStatus = "ֵ�����": thisAMStatus = thisPMStatus
                            ElseIf (thisAMStatus = "�¼�" And thisPMStatus = "�¼�") Or (thisAMStatus = "�¼�" And thisPMStatus = "�¼�") Then
                                     thisPMStatus = "ֵ�����": thisAMStatus = thisPMStatus
                            End If
                        End If
                    End If
                Else
                    If weekDay <> "����" And weekDay <> "����" And Not isFday Then
                        If tmpName = "������" Or tmpName = "������" Or tmpName = "�Ծ�" Or tmpName = "������" Then
                             thisPMStatus = "��ְ": thisAMStatus = thisPMStatus '1.1-2.20
                             'thisPMStatus = "���": thisAMStatus = thisPMStatus '1.1-2.20
                        ElseIf tmpName = "ɳ����" Then
                             thisPMStatus = "�¼�": thisAMStatus = thisPMStatus
                             'thisPMStatus = "�¼�": thisAMStatus = thisPMStatus
                        End If
                    Else
                        thisPMStatus = "��": thisAMStatus = thisPMStatus
                    End If
                End If
                Range(tmpColumn & I).Select
                Range(tmpColumn & I).Value = thisAMStatus
                Selection.Offset(0, 1).Select '����1����Ԫ��
                Range(Selection.Address).Value = thisPMStatus
                thisAMStatus = thisPMStatus = ""
                Selection.Offset(0, -1).Select '����1����Ԫ��
                
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
        Range("A2").Value = "2025��1��" & sheetNumber & Range(tmpColumn & "5").Value & "�ſ���ͳ�Ʊ�"
        'Sheets("����").Select
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
                Selection.Offset(0, 1).Select '����1����Ԫ��
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
        comBStrSplit = Split(Sheets("H").Range("R" & Hline).Value, "-") 'Range("X" ��������� 1R  2S 3T 4U 5V 6W 7X 8Y 9Z 10AA 11AB 12AC
        weekDay = Sheets("H").Range("N" & 226 + I).Value '��������� ���ܵ�һ����һ�У�Ҫ+i
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
        Selection.Value = "2025��1��" & I & "�գ�" & weekDay & "������ͳ�Ʊ�"
    Next
    'ActiveWorkbook.Save
End Function

Function clearContent()
    Sheets("����-��ֵ").Select
    'recordLine = 350
    totalLine = 299
    LogNum = 1
    For Line = 7 To totalLine
        cellName = "CJ" & Line
        For Colu = 1 To 62
            Range(cellName).Select
            Selection.Offset(0, 1).Select '����1����Ԫ��
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
    'ActiveSheet.Range("$A$7:$DB$220").AutoFilter Field:=6, Criteria1:="��"
    Range("B8:C8").Select
    Range("B7").Select
    ActiveCell.Offset(1, 0).Select
End Function

Function testDealNewD()
    Call dealNewD(21, "����")
End Function


Function dealNewD(dayNo As Integer, weekDay As String)
    If InStr(weekDay, "��") > 0 Or InStr(weekDay, "��") > 0 Then
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
    Selection.Replace What:="-����", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-�ٵ�", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-�¼�", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-����", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Dim monthDay  As Variant
    Dim tmpStr As String
    monthDay = Array("E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T") '�У���������
    For I = 79 To 97 '���������ϵ���
        For j = 0 To 15 '�У���������
            tmpStr = Range(monthDay(j) & I).Value
            If Len(tmpStr) > 1 Then
                Debug.Print tmpStr
                Range(monthDay(j) & I - 24).Select
                If InStr(tmpStr, Chr(10)) > 0 Then
                    Dim names As Variant
                    names = Split(tmpStr, Chr(10)) '����״̬
                    Dim upBound As Integer
                    Dim totalValue As Single
                    upBound = UBound(names)
                    For k = 0 To upBound
                        Debug.Print names(k)
                        If Not InStr(names(k), "��") > 0 Then
                            totalValue = totalValue + CSng(KeepNumbersAndDecimals((names(k))))
                            Debug.Print KeepNumbersAndDecimals((names(k)))
                            Debug.Print totalValue
                        End If
                    Next
                    Range(monthDay(j) & I - 24).Value = totalValue
                ElseIf Not InStr(tmpStr, "��") > 0 Then
                    Range(monthDay(j) & I - 24).Value = KeepNumbersAndDecimals(tmpStr) '79-55=24
                End If
            Else
                Range(monthDay(j) & I - 24).Value = ""
            End If
            totalValue = 0
        Next
    Next
    
    ChDir "D:\����\�ŶӺ�\����"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "D:\����\�ŶӺ�\����\1��" & dayNo & "�տ����ձ�.pdf", Quality:=xlQualityStandard, IncludeDocProperties _
        :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Function


Function dayReportBook(dayNo As Integer) As String '�����ձ����

    Dim columnNumber, Line, StartPosi, EndPosi As Integer
    Dim posiPre, status, firstValue, secValue, thisName  As String
    Dim repeatSign As Boolean
    Dim posi, tempStatus, tmpWeekDay As String
    Dim tempRepeatStatus As Variant '������״̬
    totalLine = 176
    columnNumber = dayNo
    
    Sheets("day").Select
    Range("E6:H23").ClearContents 'Ӧ���ڡ�ʵ����
    Range("I6:BC23").ClearContents '���Ű����~��ְ
    Range("BE6").Value = "" '�ظ�
    Range("BG6:BJ23").ClearContents '����һ��ĳ�����������ֵ�ࡢ�Ӱ�
    Range("BE24").ClearContents
    
    Columns("A:BE").Select
    Range("BE3").Activate
    Selection.EntireColumn.Hidden = False
    Range("O5").Select
    
    For I = 6 To 23 'ȡӦ��������
        Range("E" & I).Select
        Range("E" & I).Value = Sheets("H").Range(GetColumnName(5 + columnNumber) & (510 + I - 5)).Value
    Next
    
    Sheets("����").Select
    
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
        cellName = posiPre & Line '1��ǰ��һ��
        Sheets("����").Select
        Range(cellName).Select
        Selection.Offset(0, 1).Select '����1����Ԫ�� ����
        cellName = Selection.Address
        firstValue = Range(Selection.Address).Value '����
        Selection.Offset(0, 1).Select '����1����Ԫ��
        secValue = Range(Selection.Address).Value
        Selection.Offset(0, -1).Select '����1����Ԫ�� ���ƻ�����
        ColumnName = GetColumnName(Selection.Column) '��������
        thisName = Range("C" & Selection.Row).Value
        If thisName = "������" Then
            If Not (firstValue = "��" And secValue = "��") Then
                Sheets("day").Range("E8").Value = Sheets("day").Range("E8").Value + 1
            End If
        End If
        Debug.Print "����--" & thisName & "--�к�--" & ColumnName & Selection.Row & "+++" & firstValue & "-" & secValue
        If firstValue = "����" And secValue = "����" Then
            status = "����"
        ElseIf firstValue = "��" And secValue = "��" Then
            status = "������"
        ElseIf firstValue = "�ٵ�1��" Then
            status = "�ٵ�1��"
        ElseIf firstValue = "����1��" Then
            status = "����1��"
        ElseIf firstValue = "����" And secValue = "����" Then
            status = "����"
        ElseIf firstValue = "ֵ��" And secValue = "ֵ��" Then
            status = "ֵ��"
        ElseIf firstValue = secValue Then
            status = firstValue
        ElseIf InStr(firstValue, Chr(10)) > 0 Then ' ����˫��״̬
            'callDebug
            Debug.Print "����--" & thisName & "--�к�--" & ColumnName & Selection.Row & "+++" & firstValue & "-" & secValue
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
        ElseIf InStr(secValue, Chr(10)) > 0 Then ' ����˫��״̬
            Debug.Print "����--" & thisName & "--�к�--" & ColumnName & Selection.Row & "+++" & firstValue & "-" & secValue
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
        ElseIf firstValue <> "��" And (secValue = "��" Or secValue = "") Then
            status = statusDeal((firstValue))
        ElseIf secValue <> "��" And (firstValue = "��" Or firstValue = "") Then
            status = statusDeal((secValue))
            'If firstValue = "����" Or firstValue = "��" Or firstValue = "" Then
            '    status = secValue
            'End If
        ElseIf firstValue = "�ٵ�" And secValue = "����" Then
            status = "�ٵ�"
        ElseIf firstValue = "δǩ��" And secValue = "����" Then
            status = "δǩ��"
        ElseIf firstValue = "����" And secValue = "δǩ��" Then
            status = "δǩ��"
        ElseIf firstValue = "����" And secValue = "δ��" Then
            status = "δǩ��"
        ElseIf firstValue = "��1" And (secValue = "����" Or secValue = "ֵ��") Then
            'status = statusDeal((firstValue))
            status = "�ٵ�0.125��"
        ElseIf firstValue = "��4" And (secValue = "����" Or secValue = "ֵ��") Then
            'status = statusDeal((firstValue))
            status = "�ٵ�0.5��"
        ElseIf (firstValue = "����" Or firstValue = "ֵ��") And secValue = "��1" Then
            status = "����0.125��"
        ElseIf (firstValue = "����" Or firstValue = "ֵ��") And secValue = "��4" Then
            status = "����0.5��"
        ElseIf firstValue = "����" Then
            'callDebug
            status = statusDeal((secValue))
        Else
            If firstValue <> secValue Then
                'callDebug
                status = statusDeal((firstValue))
                If secValue <> "����" Then
                    'callDebug
                    notRepeat = False
                End If
            Else
                callDebug
            End If
        End If
        If status <> "" Then Call dealDayReport((Line), (status))
        If notRepeat = False Then ' �����ȴ���ȫ��ĵ���״̬������������������ͬ���ٴ��������
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
    
    For I = 1 To 23 '1 = k �¼� 11
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
    Range("A2").Value = "2025��1��" & columnNumber & "�ſ���ͳ�Ʊ�"
    Sheets("d").Select
    '�й�����Դ2024��  ��  �գ�����   ��ȫԱ���ڹ�ʾ��ԭʼ��
    Range("A2").Value = "�й�����Դ2025��1��" & dayNo & "�գ�����" & Right(tmpWeekDay, 1) & "��ȫԱ���ڹ�ʾ��ԭʼ��"
    Sheets("d7").Select
    '�й�����Դ2024��  ��  �գ�����   ��ȫԱ���ڹ�ʾ��ԭʼ��
    Range("A2").Value = "�й�����Դ2025��1��" & dayNo & "�գ�����" & Right(tmpWeekDay, 1) & "��ȫԱ���ڹ�ʾ��ԭʼ��"
    Sheets("day").Select
    'ActiveWorkbook.Save
End Function

Function doublePerLine()   '��Ԫ����ÿ����ʾ2�ˣ��Լ��ٵ�Ԫ��߶� H���� R��� APֵ�� AZ������
   
     Dim arr, arrType As Variant
     Dim str As String
     Dim tmpCount As Integer

    '����Ҫ���������
    arrType = Array("H", "R", "AP", "AZ")
    
    Sheets("day").Select
    Range("D5").Select
    'For m = 6 To 23
    For m = 6 To 23
        For j = 0 To 3 ' H���� R��� APֵ�� AZ������
            Range(arrType(j) & m).Select
            arr = Split(Range(arrType(j) & m).Value, Chr(10))
            'Debug.Print Range(arrType(j) & m).Value
            'Debug.Print LBound(arr)
            'Debug.Print UBound(arr)
            '��������Ľ��
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
                        str = str & "��" & arr(ii)
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
    If InStr(str, "�Ӱ�") > 0 Then
        tmpStr = "�Ӱ�"
    ElseIf InStr(str, "��") > 0 Then
        tmpStr = "�Ӱ�" & Right(Left(str, 2), 1) & "Сʱ"
    ElseIf InStr(str, "ֵ��") > 0 Then
        tmpStr = "ֵ��"
    ElseIf InStr(str, "ֵ") > 0 Then
        tmpStr = "ֵ��" & Right(Left(str, 2), 1) & "Сʱ"
    ElseIf InStr(firstValue, "����") > 0 Then
        tmpStr = "����"
    ElseIf InStr(firstValue, "��") > 0 Then
        tmpStr = "����" & Right(Left(str, 2), 1) & "Сʱ"
    ElseIf InStr(str, "����") > 0 Then
        tmpStr = "����"
    ElseIf InStr(str, "��") > 0 Then
        callDebug
        tmpStr = "����" & Right(Left(str, 2), 1) & "Сʱ"
    ElseIf InStr(str, "�¼�") > 0 Then
        tmpStr = "�¼�"
    ElseIf InStr(str, "��") > 0 Then
        tmpStr = "�¼�" & Right(Left(str, 2), 1) & "Сʱ"
    ElseIf InStr(str, "����") > 0 Then
        tmpStr = "����"
    ElseIf InStr(str, "��") > 0 Then
        tmpStr = "����" & Right(Left(str, 2), 1) & "Сʱ"
    ElseIf InStr(str, "����") > 0 Then
        tmpStr = "����"
    ElseIf InStr(str, "����") > 0 Then
        tmpStr = "����"
    ElseIf InStr(str, "��") > 0 Then
        tmpStr = "����" & Right(Left(str, 2), 1) & "Сʱ"
    Else
        tmpStr = str
    End If
    statusDeal = tmpStr
    If InStr(tmpStr, "��Сʱ") > 0 Then callDebug
End Function

Function dealDayReport(Line As Integer, thisStatus As String)

'�¼� i  ���� k   ���� m   ���˼� O   ��� Q  ���� S  ��� U  ɥ�� W  ���� Y  �ٵ� AA  ���� AC δǩ�� AE   δǩ�� AG   ��ǩ�� AI
'��ǩ�� AK       ���� AM     ֵ�� AO     ֵ����� AQ     ֵ����� AS     �Ӱ� AU     ���� AW     ������ AY       ��ְ BA

'4   CHYY�ܲ� '5   ���ܲ�ҵ '6   ���ӳ�Ա '7   �ۺ����� '8   ���� '9   ���ݹ�˾ '10  ����Դ1��ҵ�� '11  ����Դ2��ҵ�� '12  ����Դ3��ҵ��
'13  ����Դ4��ҵ�� '14  �ȱ÷ֹ�˾ '15  ԴȪ�ֹ�˾ '16  ��װ��˾ '17  ������Դ��˾ '18  ������ҵ�� '19  �������乫˾ '20  ��ά1��ҵ�� '21  ��ά2��ҵ��
    
        Dim coName(), StatusName() As Variant, tempCoName, rangePosi, employerName, status As String, flagRow, flagCol As Integer
        Dim notAlldayColumn As String
        Sheets("����").Select
        tempCoName = Range("F" & Line).Value 'Ŀǰ�� ���� ��
        employerName = Range("C" & Line).Value 'Ŀǰ�� ���� ��
        status = thisStatus
        Debug.Print status
        Sheets("day").Select
        coName() = Range("A6:A23").Value
        StatusName() = Range("H5:BC5").Value
        'status = "����3Сʱ"
        'status = "����"
        flagRow = WorksheetFunction.Match(tempCoName, coName, 0) + 5 '������5��
        Debug.Print flagRow
        If status = "���" Then status = "����"
        If InStr(status, "Сʱ") > 0 Or InStr(status, "��") > 0 Then
            flagCol = WorksheetFunction.Match(Left(status, 2), StatusName, 0) + 7 'ǰ����7��
        Else
            flagCol = WorksheetFunction.Match(status, StatusName, 0) + 7 'ǰ����7��
        End If
        rangePosi = GetColumnName(flagCol) & flagRow
        
        Debug.Print rangePosi
        
        Range(rangePosi).Select
        'If status = "����" Then ' Or status = "������"
            '��
        'ElseIf Range(rangePosi).Value = "" Then
        If Range(rangePosi).Value = "" Then
            If InStr(status, "Сʱ") > 0 Or InStr(status, "��") > 0 Then
                Range(rangePosi).Value = employerName & "-" & status
                If Not InStr(status, "��") > 0 Then
                    notAlldayColumn = obtainNotAllDayColumn(status)
                    Range(notAlldayColumn & flagRow).Value = Range(notAlldayColumn & flagRow).Value + 1
                End If
            Else
                Range(rangePosi).Value = employerName
            End If
        Else
            If InStr(status, "Сʱ") > 0 Or InStr(status, "��") > 0 Then
                Range(rangePosi).Value = Range(rangePosi).Value & Chr(10) & employerName & "-" & status
                If Not InStr(status, "��") > 0 Then
                    notAlldayColumn = obtainNotAllDayColumn(status)
                    Range(notAlldayColumn & flagRow).Value = Range(notAlldayColumn & flagRow).Value + 1
                End If
            Else
                Range(rangePosi).Value = Range(rangePosi).Value & Chr(10) & employerName
            End If
        End If
        Selection.Offset(0, 1).Select '����1����Ԫ��
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
        Selection.Offset(0, -1).Select '����1����Ԫ��
End Function

Function obtainNotAllDayColumn(status As String) As String
    If InStr(status, "��") > 0 Then
        obtainNotAllDayColumn = "BG"
    ElseIf InStr(status, "��") > 0 Then
        obtainNotAllDayColumn = "BH"
    ElseIf InStr(status, "��") > 0 Then
        obtainNotAllDayColumn = "BI"
    ElseIf InStr(status, "��") > 0 Then
        obtainNotAllDayColumn = "BJ"
    ElseIf InStr(status, "��") > 0 And Len(status) > 2 Then
        obtainNotAllDayColumn = "BK"
    ElseIf InStr(status, "��") > 0 And Len(status) > 2 Then
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
        For I = 32 To 49 '���������ʾ��,  ������ 324-341  324-32=292
            For j = 15 To 37 'O to AK ��
            'For j = 31 To 31 '��ǩ��
                tempCol = GetColumnName(j)
                rangePosi = tempCol & (I + 292)
                Range(rangePosi).Select
                thisValue = Range(rangePosi).Value
                Debug.Print thisValue
                If Len(thisValue) > 1 Then
                    Range(tempCol & I).Select
                    If tempCol = "X" Or tempCol = "Y" Then ' �ٵ� ����
                        Range(tempCol & I).Value = getSortValueResultEarlyOrLate(thisValue)
                    ElseIf tempCol = "Z" Or tempCol = "AA" Then ' δǩ�� δǩ��
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
    names = Split(str, Chr(10)) '����״̬
    Dim upBound, tmpCont As Integer
    upBound = UBound(names)
    For k = 0 To upBound
        Debug.Print names(k)
        Dim hasThisName As Boolean
        hasThisName = False
        If k <> 0 Then '��һ��û��ͳ�ƹ����Թ�
            For l = 0 To k - 1 '
                If names(k) = names(l) Then '��ǰ״̬��ǰ���أ��Ƿ���ͳ�ƹ�������ͳ�ƹ�������ͳ��
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
            
            If Len(valueResult) < 2 Then '��һ��
                If tmpCont > 1 Then
                    If InStr(names(k), "-") > 0 Then
                        valueResult = names(k) & "-" & tmpCont & "��"
                    Else
                        valueResult = names(k) & "-" & tmpCont & "��"
                    End If
                ElseIf k = 0 Then
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = names(k) & "-1��"
                    Else
                        valueResult = names(k)
                    End If
                Else
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-1��"
                    Else
                        valueResult = valueResult & Chr(10) & names(k)
                    End If
                End If
            Else
                If tmpCont > 1 Then
                    If InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-" & tmpCont & "��"
                    Else
                        valueResult = valueResult & Chr(10) & names(k) & "-" & tmpCont & "��"
                    End If
                Else
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-1��"
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
    names = Split(str, Chr(10)) '����״̬
    Dim upBound, tmpCont As Integer
    upBound = UBound(names)
    For k = 0 To upBound
        Debug.Print names(k)
        Dim hasThisName As Boolean
        hasThisName = False
        If k <> 0 Then '��һ��û��ͳ�ƹ����Թ�
            For l = 0 To k - 1 '
                If names(k) = names(l) Then '��ǰ״̬��ǰ���أ��Ƿ���ͳ�ƹ�������ͳ�ƹ�������ͳ��
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
            
            If Len(valueResult) < 2 Then '��һ��
                If tmpCont > 1 Then
                    valueResult = names(k) & "-" & tmpCont & "��"
                ElseIf k = 0 Then
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = names(k) & "-1��"
                    Else
                        valueResult = names(k)
                    End If
                Else
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-1��"
                    Else
                        valueResult = valueResult & Chr(10) & names(k)
                    End If
                End If
            Else
                If tmpCont > 1 Then
                    valueResult = valueResult & Chr(10) & names(k) & "-" & tmpCont & "��"
                Else
                    If Not InStr(names(k), "-") > 0 Then
                        valueResult = valueResult & Chr(10) & names(k) & "-1��"
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
    names = Split(str, Chr(10)) '����״̬
    Dim upBound, tmpCont As Integer
    upBound = UBound(names)
    For k = 0 To upBound
        Debug.Print names(k)
        Dim hasThisName As Boolean
        hasThisName = False
        If k <> 0 Then '��һ��û��ͳ�ƹ����Թ�
            For l = 0 To k - 1 '
                If names(k) = names(l) Then '��ǰ״̬��ǰ���أ��Ƿ���ͳ�ƹ�������ͳ�ƹ�������ͳ��
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
            
            If Len(valueResult) < 2 Then '��һ��
                If tmpCont > 1 Then
                    valueResult = names(k) & "-" & tmpCont * 0.5 & "��"
                ElseIf k = 0 Then
                    valueResult = names(k) & "-0.5��"
                Else
                    valueResult = valueResult & Chr(10) & names(k)
                End If
            Else
                If tmpCont > 1 Then
                    valueResult = valueResult & Chr(10) & names(k) & "-" & tmpCont * 0.5 & "��"
                Else
                    valueResult = valueResult & Chr(10) & names(k) & "-0.5��"
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
        For I = 15 To 37 ' �� �¼� ���������31ֵ��
            thisCol = GetColumnName(I)
            firValue = UBound(Split(Range(thisCol & thisRow).Value, Chr(10))) + 1
            If I = 31 Or I = 36 Then
                If firValue > 6 Then
                    'Range(thisCol & thisRow).Value = "��"
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
    'ap8 ֵ�� az8 �� 7~23
'    For i = 7 To 23 ' ֵ�� �� ������
'        Range("ap" & i).Select
'        lineNumber = UBound(Split(Range("ap" & i).Value, Chr(10))) + 1
'        If lineNumber > 6 Then
'            'Range("ap" & i).Value = "��"
'            'Range("BL" & i).Value = 2
'        End If
'
'        Range("az" & i).Select
'        lineNumber = UBound(Split(Range("az" & i).Value, Chr(10))) + 1
'        If lineNumber > 6 Then
'            'Range("az" & i).Value = "��"
'            'Range("BL" & i).Value = 2
'        End If
'    Next
'
'    For i = 6 To 23 '���� ���������ж�ÿһ���ڵ����������������ֵȡ�и�
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
    For I = 1 To 23 '1 = k �¼� 11 '0ֵ����
        thisCol = GetColumnName(2 * I - 1 + 10)
        Range(thisCol & "24").Select
        preCol = GetColumnName(2 * I - 1 + 9)
        If Range(thisCol & "24").Value = 0 Then
            Debug.Print preCol & "-" & thisCol
            Columns(preCol).Hidden = True
            Columns(thisCol).Hidden = True
        Else
            For j = 6 To 23 '����
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
    For I = 1 To 23 '1 = k �¼� 11 '0ֵ���� 16 ����
        thisCol = GetColumnName(14 + I)
        Debug.Print "WeekColumnsWidth-" & thisCol
        Range(thisCol & "24").Select
        If Range(thisCol & "24").Value = 0 Then
            Columns(thisCol).Hidden = True
        Else
            For j = 32 To 49 '����
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
    For I = 15 To 37 ' E �¼� AK
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
    Sheets("����").Select
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
        If Range(posi & dataStartLine + I).Value > 0 And I <> 10 Then '��ְ��
            lineNum = lineNum + 1
            Range(posiPre & 318 + I).Value = lineNum & Chr(10) & Range(posiPre & 318 + I).Value
            tmpSum = tmpSum + Range(posi & dataStartLine + I).Value
        End If
    Next I
    For Line = 7 To totalLine + 6
        cellName = posiPre & Line '1��ǰ��һ��
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
    Selection.Offset(0, 1).Select '����1����Ԫ��
    cellName = Selection.Address
    ColumnName = GetColumnName(Selection.Column)
    Debug.Print "����--" & Range("C" & Selection.Row).Value & "--�к�--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
    'If IsNumeric(Range(Selection.Address).Value) Then
    '    Range(posi & posiRow).Value = Range(posi & posiRow).Value & "��" & Range("C" & Selection.Row).Value & "-" & Range(Selection.Address).Value & "����" 'ȡ������ֵ
    '    Debug.Print Range(posi & posiRow).Value
    If Range(Selection.Address).Value = "�ٵ�" Then
        Range(posi & posiRow).Value = Range(posi & posiRow).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
        Debug.Print Range(posi & posiRow).Value
    ElseIf Range(Selection.Address).Value = "ֵ�����" Then
        Range(posi & posiRow + 20 + 1).Value = Range(posi & posiRow + 20 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
        Debug.Print Range(posi & posiRow + 20 + 1).Value
        Range(posi & posiRow + 20 + 1).Select
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        If Range(Selection.Address).Value = "����" Then
            Range(posi & posiRow + 1 + 1).Value = Range(posi & posiRow + 1 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
        Else
            Range(posi & posiRow + 1 + 1).Value = Range(posi & posiRow + 1 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ" 'ȡ������ֵ
        End If
        Debug.Print Range(posi & posiRow + 1 + 1).Value
    ElseIf Range(Selection.Address).Value = "����" Then
        Range(posi & posiRow + 2 + 1).Value = Range(posi & posiRow + 2 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
        Debug.Print Range(posi & posiRow + 2 + 1).Value
    ElseIf Range(Selection.Address).Value = "����" Then
        Range(posi & posiRow + 3 + 1).Value = Range(posi & posiRow + 3 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
        Debug.Print Range(posi & posiRow + 3 + 1).Value
    ElseIf Range(Selection.Address).Value = "���" Then
        Range(posi & posiRow + 10 + 1).Value = Range(posi & posiRow + 10 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
        Debug.Print Range(posi & posiRow + 10 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        Range(posi & posiRow + 10 + 1).Value = Range(posi & posiRow + 10 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ" 'ȡ������ֵ
        Debug.Print Range(posi & posiRow + 10 + 1).Value
    ElseIf Range(Selection.Address).Value = "��" Then
        Range(posi & posiRow + 4 + 1).Value = Range(posi & posiRow + 4 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
        Debug.Print Range(posi & posiRow + 4 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        If Range(Selection.Address).Value = "�¼�" Then
            Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
        Else
            Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ" 'ȡ������ֵ
        End If
        Debug.Print Range(posi & posiRow + 5 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        If Range(Selection.Address).Value = "����" Then
            Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
        Else
            Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ"  'ȡ������ֵ
        End If
        Debug.Print Range(posi & posiRow + 6 + 1).Value
    ElseIf Range(Selection.Address).Value = "δǩ��" Then
        Range(posi & posiRow + 7 + 1).Value = Range(posi & posiRow + 7 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
        Debug.Print Range(posi & posiRow + 7 + 1).Value
    ElseIf Range(Selection.Address).Value = "��ǩ��" Then
        Range(posi & posiRow + 14 + 1).Value = Range(posi & posiRow + 14 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
        Debug.Print Range(posi & posiRow + 14 + 1).Value
'    ElseIf Range(Selection.Address).Value = "����" Then
'        Range(posi & posiRow + 14 + 1).Value = Range(posi & posiRow + 14 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
        Debug.Print Range(posi & posiRow + 14 + 1).Value
    ElseIf Range(Selection.Address).Value = "��ְ" Then
        Range(posi & posiRow + 8 + 1).Value = Range(posi & posiRow + 8 + 1).Value & "��" & Range("C" & Selection.Row).Value    'ȡ����
        Debug.Print Range(posi & posiRow + 8 + 1).Value
    ElseIf Range(Selection.Address).Value = "��ְ" Then
        Range(posi & posiRow + 9 + 1).Value = Range(posi & posiRow + 9 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
        Debug.Print Range(posi & posiRow + 9 + 1).Value
    ElseIf Range(Selection.Address).Value = "����" Then
        Range(posi & posiRow + 12 + 1).Value = Range(posi & posiRow + 12 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
        Debug.Print Range(posi & posiRow + 12 + 1).Value
    ElseIf Range(Selection.Address).Value = "���" Then
        Range(posi & posiRow + 13 + 1).Value = Range(posi & posiRow + 13 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
        Debug.Print Range(posi & posiRow + 13 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "��") Then '����1 2 3
        Range(posi & posiRow + 3 + 1).Value = Range(posi & posiRow + 3 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ" 'ȡ������ֵ
        Debug.Print Range(posi & posiRow + 3 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "����") Then '
        Range(posi & posiRow + 12 + 1).Value = Range(posi & posiRow + 12 + 1).Value & "��" & Range("C" & Selection.Row).Value    'ȡ����
        Debug.Print Range(posi & posiRow + 12 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "����") Then '
        Range(posi & posiRow + 14 + 1).Value = Range(posi & posiRow + 14 + 1).Value & "��" & Range("C" & Selection.Row).Value    'ȡ����
        Debug.Print Range(posi & posiRow + 14 + 1).Value
    ElseIf InStr(Range(Selection.Address).Value, "ɥ��") Then '
        Range(posi & posiRow + 15 + 1).Value = Range(posi & posiRow + 15 + 1).Value & "��" & Range("C" & Selection.Row).Value    'ȡ����
        Debug.Print Range(posi & posiRow + 15 + 1).Value
'    ElseIf InStr(Range(Selection.Address).Value, "����") Then '
'        Range(posi & posiRow + 17 + 1).Value = Range(posi & posiRow + 17 + 1).Value & "��" & Range("C" & Selection.Row).Value    'ȡ����
'        Debug.Print Range(posi & posiRow + 17 + 1).Value
    End If
    
    Selection.Offset(0, 1).Select '����1����Ԫ��
    cellName = Selection.Address
    Debug.Print "next cell " & cellName
    ColumnName = GetColumnName(Selection.Column)
    Debug.Print "����--" & Range("C" & Selection.Row).Value & "--�к�--" & ColumnName & Selection.Row & "===" & Range(Selection.Address).Value
    
    If IsNumeric(Range(Selection.Address).Value) Then
        If Not InStr(Range(posi & posiRow + 1).Value, Range("C" & Selection.Row).Value) Then
            Range(posi & posiRow + 1).Value = Range(posi & posiRow + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & -Range(Selection.Address).Value & "����"   'ȡ������ֵ
            Debug.Print Range(posi & posiRow + 1).Value
            Range(posi & posiRow + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "ֵ�����" Then
        If Not InStr(Range(posi & posiRow + 20 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 20 + 1).Value = Range(posi & posiRow + 20 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
            Debug.Print Range(posi & posiRow + 20 + 1).Value
            Range(posi & posiRow + 20 + 1).Select
        End If
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        If Not InStr(Range(posi & posiRow + 1 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            If Range(Selection.Address).Value = "����" Then
                Range(posi & posiRow + 1 + 1).Value = Range(posi & posiRow + 1 + 1).Value & "��" & Range("C" & Selection.Row).Value    'ȡ����
            Else
                Range(posi & posiRow + 1 + 1).Value = Range(posi & posiRow + 1 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ" 'ȡ������ֵ
            End If
            Debug.Print Range(posi & posiRow + 1).Value
            Range(posi & posiRow + 1 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "����" Then
        If Not InStr(Range(posi & posiRow + 3 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 3 + 1).Value = Range(posi & posiRow + 3 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
            Debug.Print Range(posi & posiRow + 3).Value
            Range(posi & posiRow + 3 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "���" Then
        If Not InStr(Range(posi & posiRow + 10 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 10 + 1).Value = Range(posi & posiRow + 10 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
            Debug.Print Range(posi & posiRow + 10 + 1).Value
            Range(posi & posiRow + 10 + 1).Select
        End If
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        If Not InStr(Range(posi & posiRow + 10 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 10 + 1).Value = Range(posi & posiRow + 10 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1)  'ȡ������ֵ
            Debug.Print Range(posi & posiRow + 10 + 1).Value
            Range(posi & posiRow + 10 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "��" Then
        If Not InStr(Range(posi & posiRow + 4 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 4 + 1).Value = Range(posi & posiRow + 4 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
            Debug.Print Range(posi & posiRow + 4 + 1).Value
            Range(posi & posiRow + 4 + 1).Select
        End If
'    ElseIf Range(Selection.Address).Value = "����" Then
'        If Not InStr(Range(posi & posiRow + 14 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
'            Range(posi & posiRow + 14 + 1).Value = Range(posi & posiRow + 14 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
'            Debug.Print Range(posi & posiRow + 14 + 1).Value
'            Range(posi & posiRow + 14 + 1).Select
'        End If
    ElseIf Range(Selection.Address).Value = "���" Then 'ԭ����
        If Not InStr(Range(posi & posiRow + 13 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 13 + 1).Value = Range(posi & posiRow + 13 + 1).Value & "��" & Range("C" & Selection.Row).Value    'ȡ����
            Debug.Print Range(posi & posiRow + 13 + 1).Value
            Range(posi & posiRow + 13 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "δǩ��" Then
        If Not InStr(Range(posi & posiRow + 16 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 16 + 1).Value = Range(posi & posiRow + 16 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
            Debug.Print Range(posi & posiRow + 16 + 1).Value
            Range(posi & posiRow + 16 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "��ǩ��" Then
        If Not InStr(Range(posi & posiRow + 17 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 17 + 1).Value = Range(posi & posiRow + 17 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
            Debug.Print Range(posi & posiRow + 17 + 1).Value
            Range(posi & posiRow + 17 + 1).Select
        End If
'    ElseIf Range(Selection.Address).Value = "����" Then
'        If Not InStr(Range(posi & posiRow + 17 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
'            Range(posi & posiRow + 17 + 1).Value = Range(posi & posiRow + 17 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
'            Debug.Print Range(posi & posiRow + 17 + 1).Value
'            Range(posi & posiRow + 17 + 1).Select
'        End If
    ElseIf Range(Selection.Address).Value = "����" Then
        If Not InStr(Range(posi & posiRow + 18 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 18 + 1).Value = Range(posi & posiRow + 18 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
            Debug.Print Range(posi & posiRow + 18 + 1).Value
            Range(posi & posiRow + 18 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "���˼�" Then
        If Not InStr(Range(posi & posiRow + 19 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 19 + 1).Value = Range(posi & posiRow + 19 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
            Debug.Print Range(posi & posiRow + 19 + 1).Value
            Range(posi & posiRow + 19 + 1).Select
        End If
    ElseIf Range(Selection.Address).Value = "ɥ��" Then
        If Not InStr(Range(posi & posiRow + 15 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 15 + 1).Value = Range(posi & posiRow + 15 + 1).Value & "��" & Range("C" & Selection.Row).Value   'ȡ����
            Debug.Print Range(posi & posiRow + 15 + 1).Value
            Range(posi & posiRow + 15 + 1).Select
        End If
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        If Not InStr(Range(posi & posiRow + 5 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ" 'ȡ������ֵ
            Range(posi & posiRow + 5 + 1).Select
'        Else
'            If Range(Selection.Address).Value = "�¼�" Then
'                Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
'            Else
'                Range(posi & posiRow + 5 + 1).Value = Range(posi & posiRow + 5 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ" 'ȡ������ֵ
'            End If
        End If
        Debug.Print Range(posi & posiRow + 5 + 1).Value
        
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        If Not InStr(Range(posi & posiRow + 6 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1)  'ȡ������ֵ
            Debug.Print Range(posi & posiRow + 6 + 1).Value
            Range(posi & posiRow + 6 + 1).Select
'        Else
'            If Range(Selection.Address).Value = "����" Then
'                Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "��" & Range("C" & Selection.Row).Value  'ȡ����
'            Else
'                Range(posi & posiRow + 6 + 1).Value = Range(posi & posiRow + 6 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ"  'ȡ������ֵ
'            End If
        End If
    ElseIf InStr(Range(Selection.Address).Value, "��") And Len(Range(Selection.Address).Value) > 1 Then
        If Not InStr(Range(posi & posiRow + 3 + 1).Value, Range("C" & Selection.Row).Value) > 0 Then
            Range(posi & posiRow + 3 + 1).Value = Range(posi & posiRow + 3 + 1).Value & "��" & Range("C" & Selection.Row).Value & "-" & Right(Range(Selection.Address).Value, 1) & "Сʱ" 'ȡ������ֵ
            Debug.Print Range(posi & posiRow + 3).Value
            Range(posi & posiRow + 3 + 1).Select
        End If
    End If
End Function

Function dealBlank() As String
    Sheets("����").Select
    For I = 7 To 210
        Range("H" & I).Select
        For j = 1 To 50
            Selection.Offset(0, 1).Select '����1����Ԫ��
            Debug.Print Range("K8").Value
            cellName = Selection.Address
            ColumnName = GetColumnName(Selection.Column)
            Debug.Print "����--" & Range("C" & Selection.Row).Value & "--�к�--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
            If Trim(Range(Selection.Address).Value) = "" Or Trim(Range(Selection.Address).Value) = 0 Then
                Range(Selection.Address).Value = "��"
            End If
        Next j
    Next I
End Function

Function justSecquence()
    Dim tmpName As String
    
    For I = 7 To 176
        Sheets("����").Select
        tmpName = Range("C" & I).Value
        Debug.Print tmpName
        Sheets("W1").Select
        If Range("E" & I).Value <> tmpName Then
            Cells.Find(What:=tmpName, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                , MatchByte:=False, SearchFormat:=False).Activate
            tmpRow = Selection.Row
            Debug.Print tmpName & ">����������>" & tmpRow
            Rows(tmpRow & ":" & tmpRow).Select
            Selection.Cut
            Rows(I & ":" & I).Select
            Selection.Insert Shift:=xlDown
        End If
        Delayms (1)
    Next
End Function

Function ll()
    Sheets("����").Select
    Dim posi As String
    posi = "H"
    recordLine = 410
    LogNum = 1
    For Line = 7 To 176
        cellName = posi & Line '1��ǰ��һ��
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
    Selection.Offset(0, 1).Select '����1����Ԫ��
    cellName = Selection.Address
    ColumnName = GetColumnName(Selection.Column)
    Debug.Print "����--" & Range("C" & Selection.Row).Value & "--�к�--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
    If Range(Selection.Address).Value = "" Or Range(Selection.Address).Value = 0 Or InStr(Range(Selection.Address).Value, "����") Then
        Range(Selection.Address).Value = "��"
    ElseIf Range(Selection.Address).Value = "�ٵ�" Or (IsNumeric(Range(Selection.Address).Value) And Range(Selection.Address).Value > 0) Then
        'la (Selection.Address)
        If Range(Selection.Address).Value = "�ٵ�" Then
            Range(ColumnName & "1").Select
            Selection.Copy
            Range(cellName).Select
            ActiveSheet.Paste
            'ActiveWorkbook.Save
        End If
        If Range(Selection.Address).Value >= 30 Then
            If vStr = "" Then
                     vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                            Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "�ٵ�" & Range(Selection.Address).Value & "����"
                Debug.Print vStr
            Else
                vStr = vStr & "," & " " & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "�ٵ�" & Range(Selection.Address).Value & "����"
                Debug.Print vStr
            End If
        End If
        Debug.Print Selection.Address & "�ٵ�"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "ȱ��" Or Range(Selection.Address).Value = "δǩ��" Then
        Range(Selection.Address).Value = "δǩ��"
        If vStr = "" Then
            vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "δǩ��"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "δǩ��"
        End If
        Debug.Print Selection.Address & "δǩ��"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "��ǩ��" Then
        Range(Selection.Address).Value = "��ǩ��"
        If vStr = "" Then
            vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "��ǩ��"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "��ǩ��"
        End If
        Debug.Print Selection.Address & "��ǩ��"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "����" Then
        If vStr = "" Then
            vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                    Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & "����"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & "����"
        End If
        Debug.Print Selection.Address & "����"
        Debug.Print vStr
    End If
    
    Selection.Offset(0, 1).Select '����1����Ԫ��
    cellName = Selection.Address
    ColumnName = GetColumnName(Selection.Column)
    Debug.Print "����--" & Range("C" & Selection.Row).Value & "--�к�--" & ColumnName & Selection.Row & "===" & Range(Selection.Address).Value
    
    If Range(Selection.Address).Value = "" Or Range(Selection.Address).Value = 0 Or InStr(Range(Selection.Address).Value, "����") Then
        Range(Selection.Address).Value = "��"
    ElseIf Range(Selection.Address).Value = "����" Or (IsNumeric(Range(Selection.Address).Value) And Range(Selection.Address).Value < 0) Then
        If Range(Selection.Address).Value = "����" Then 'le (Selection.Address)
            Range(ColumnName & "1").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range(cellName).Select
            ActiveSheet.Paste
            'ActiveWorkbook.Save
        End If
        
        If Range(Selection.Address).Value <= 30 Then
            If vStr = "" Then
              vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                     Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "����" & -Range(Selection.Address).Value & "����"
                Debug.Print vStr
            Else
                vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "����" & -Range(Selection.Address).Value & "����"
                Debug.Print vSt
            End If
        End If
    ElseIf Range(Selection.Address).Value = "ȱ��" Or Range(Selection.Address).Value = "δǩ��" Then
        Range(Selection.Address).Value = "δǩ��"
        If vStr = "" Then
            vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "δǩ��"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "δǩ��"
        End If
        Debug.Print Selection.Address & "δǩ��"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "��ǩ��" Then
        Range(Selection.Address).Value = "��ǩ��"
        If vStr = "" Then
            vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "��ǩ��"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "��ǩ��"
        End If
        Debug.Print Selection.Address & "��ǩ��"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "����" Then
        Range(Selection.Address).Value = "����"
        If vStr = "" Then
            vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "����"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "����"
        End If
        Debug.Print Selection.Address & "����"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "���˼�" Then
        Range(Selection.Address).Value = "���˼�"
        If vStr = "" Then
            vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "���˼�"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "���˼�"
        End If
        Debug.Print Selection.Address & "���˼�"
        Debug.Print vStr
    ElseIf Range(Selection.Address).Value = "��δ��" Then
        Range(Selection.Address).Value = "����"
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        If vStr = "" Then
            vStr = LogNum & "�� " & Range("B" & Selection.Row).Value & " " & Range("C" & Selection.Row).Value & " " & Range("E" & Selection.Row).Value & " " & _
                   Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "��δ��"
        Else
            vStr = vStr & "," & Range(ColumnName & "4").Value & Range(ColumnName & "5").Value & Range(ColumnName & "6").Value & "��δ��"
        End If
        Debug.Print Selection.Address & "��δ��"
        Debug.Print vStr
    End If
End Function

Function dayWeekReport()
    Sheets("����").Select
    Dim posiPreLeft  As String, posiPreRight As String
    Dim posi, posiCounter As String
    hasChange = False
    cd = 0: zt = 0: kg = 0: cc = 0: dx = 0: zcx = 0: sj = 0: bj = 0: qk = 0: rz = 0: lz = 0: nx = 0: bc = 0: bt = 0: bk = 0: wqd = 0: wqt = 0: bqd = 0: bqt = 0  '�ٵ������ˡ�������������ݡ��¼١����١�ȱ�������١����ˡ�����
    cds = 0: zts = 0: kgs = 0: ccs = 0: dxs = 0: zcxs = 0: sjs = 0: bjs = 0: qks = 0: rzs = 0: lzs = 0: nxs = 0: bcs = 0: bts = 0: bks = 0: wqds = 0: wqts = 0: bqds = 0: bqts = 0 '�ٵ������ˡ�������������ݡ��¼١����١�ȱ�������١����ˡ�����
    cdStr = "":  ztStr = "":  kgStr = "":  ccStr = "":  dxStr = "": zcxStr = ""
    sjStr = "": bjStr = "": qkStr = "": rzStr = "": lzStr = "": nxStr = "": bcStr = "": btStr = "": bkStr = "": wqdStr = 0: wqtStr = 0: bqdStr = 0: bqtStr = 0
    
    posiPreLeft = "G" '�ܵ�һ�쿪ʼǰ2��'-------------------------------------------
    
    posiName = GetColumnName(GetColumnNum(posiPreLeft) + 5) '������Ϊ����
    posiCounter = GetColumnName(GetColumnNum(posiPreLeft) + 4) '������Ϊ����
    Debug.Print posiPreLeft
    Debug.Print posiName
    Debug.Print posiCounter
    
    For dayN = 1 To 7
        'For Line = 117 To 117 '
        For Line = 7 To 176 '-------------------------------------------
            cellName = GetColumnName(GetColumnNum(posiPreLeft) + (dayN - 1) * 2) & Line '1��ǰ��һ��
            Debug.Print cellName
            For Row = 1 To 1 '������
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
        
        posiPreRight = "H" '���쿪ʼǰ1��'-------------------------------------------
        'For Line = 117 To 117 '
        For Line = 7 To 176
            cellName = GetColumnName(GetColumnNum(posiPreRight) + (dayN - 1) * 2) & Line '1��ǰ��һ��
            'Debug.Print cellName
            For Row = 1 To 1 '������
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
                     Range(posiCounter & I).Value = cd & "��"
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
                     Range(posiCounter & I).Value = cd & "��"
                     Range(posiCounter & I).Value = zt & "��"
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
                     Range(posiCounter & I).Value = kg & "��"
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
                     Range(posiCounter & I).Value = cc & "��"
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
                     Range(posiCounter & I).Value = dx & "��"
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
                     Range(posiCounter & I).Value = zcx & "��"
                     Range(dateColume & I).Value = zcx
                     End If
                 Case startLine + 6
                If sjs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = sj & "��"
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
                     Range(posiCounter & I).Value = bj & "��"
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
                     Range(posiCounter & I).Value = wqd & "��"
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
                     Range(posiCounter & I).Value = rz & "��"
                     End If
                 Case startLine + 10
                If lz > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = lz & "��"
                     End If
                 Case startLine + 11
                If nxs > 0 Then
                    lineNum = lineNum + 1
                    Range(posiPreRight & I).Value = lineNum & Chr(10) & Range(posiPreRight & I).Value
                     Range(posiCounter & I).Value = nx & "��"
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
                     Range(posiCounter & I).Value = bc & "��"
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
                     Range(posiCounter & I).Value = bt & "��"
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
                     Range(posiCounter & I).Value = bqd & "��"
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
                     Range(posiCounter & I).Value = wqt & "��"
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
                     Range(posiCounter & I).Value = bqt & "��"
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
    cd = 0: zt = 0: kg = 0: cc = 0: dx = 0: zcx = 0: sj = 0: bj = 0: qk = 0: rz = 0: lz = 0: nx = 0: bc = 0: bt = 0: bqd = 0: bqt = 0: wqd = 0: wqt = 0 '�ٵ������ˡ�������������ݡ��¼١����١�ȱ�������١����ˡ�����
    cds = 0: zts = 0: kgs = 0: ccs = 0: dxs = 0: zcxs = 0: sjs = 0: bjs = 0: qks = 0: rzs = 0: lzs = 0: nxs = 0: bcs = 0: bts = 0: bqds = 0: bqts = 0: wqds = 0: wqts = 0 '�ٵ������ˡ�������������ݡ��¼١����١�ȱ�������١����ˡ�����
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
    Sheets("����").Select
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
    Sheets("��ʾ1").Select
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
    Selection.Offset(0, 2).Select '����2����Ԫ��
    cellName = Selection.Address
    ColumnName = GetColumnName(Selection.Column)
    'Debug.Print "����--" & Range("C" & Selection.Row).Value & "--�к�--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
    If Range(Selection.Address).Value = "�ٵ�" Then
        findPosi = InStr(cdStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        If findPosi <= 0 Then
            cdStr = cdStr & "��" & Range("C" & Selection.Row).Value 'ȡ���� ��һ��
            cd = cd + 1
        Else
            If Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(cdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                cdStr = Replace(cdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(cdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                cdStr = Replace(cdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(cdStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then   '�ڶ���
                cdStr = Replace(cdStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
            '1.�����ַ��� Ilovevba,�������ȡ love = Mid("Lloveba", 2, 4) ��ʼ������
        End If
        cds = cds + 1
        Debug.Print "�ٵ�=" & cdStr
    ElseIf Range(Selection.Address).Value = "�Ӱ�" Then '����
        findPosi = InStr(bcStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print bcStr
        If findPosi <= 0 Then
            'Debug.Print bcStr
            'Debug.Print Range("C" & Selection.Row).Value
            bcStr = bcStr & "��" & Range("C" & Selection.Row).Value 'ȡ���� ��һ��
            bc = bc + 1
        Else
            If Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(bcStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bcStr = Replace(bcStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(bcStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bcStr = Replace(bcStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                bcStr = Replace(bcStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & bcStr
        bcs = bcs + 1
    ElseIf Range(Selection.Address).Value = "����" Then
        findPosi = InStr(bqdStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print bkStr
        If findPosi <= 0 Then
            'Debug.Print bkStr
            'Debug.Print Range("C" & Selection.Row).Value
            bqdStr = bqdStr & "��" & Range("C" & Selection.Row).Value 'ȡ���� ��һ��
            bqd = bqd + 1
        Else
            If Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(bqdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bqdStr = Replace(bqdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(bqdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bqdStr = Replace(bqdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(bqdStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                bqdStr = Replace(bqdStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "��ǩ��=" & bqdStr
        bqds = bqds + 1
    ElseIf Range(Selection.Address).Value = "���" Then '����
        findPosi = InStr(btStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            btStr = btStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Range(posi & posiRow + 8).Select
            bt = bt + 1
        Else
            If Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(btStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                btStr = Replace(btStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(btStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                btStr = Replace(btStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                btStr = Replace(btStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & btStr
        bts = bts + 1
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        findPosi = InStr(kgStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print kgStr
        If findPosi <= 0 Then
            'Debug.Print kgStr
            'Debug.Print Range("C" & Selection.Row).Value
            kgStr = kgStr & "��" & Range("C" & Selection.Row).Value 'ȡ���� ��һ��
            kg = kg + 1
        Else
            If Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(kgStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                kgStr = Replace(kgStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(kgStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                kgStr = Replace(kgStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                kgStr = Replace(kgStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & kgStr
        kgs = kgs + 1
    ElseIf Range(Selection.Address).Value = "����" Then
        'Debug.Print "����=" & ccStr
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print Len(Range("C" & Selection.Row).Value)
        findPosi = InStr(ccStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            ccStr = ccStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            cc = cc + 1
        Else
            If Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(ccStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                ccStr = Replace(ccStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(ccStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                ccStr = Replace(ccStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(ccStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                ccStr = Replace(ccStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & ccStr
        ccs = ccs + 1
        Range(posi & posiRow + 3).Select
    ElseIf Range(Selection.Address).Value = "����" Then
        findPosi = InStr(dxStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            dxStr = dxStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            dx = dx + 1
        Else
            If Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                dxStr = Replace(dxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & dxStr
        dxs = dxs + 1
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        findPosi = InStr(nxStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            nxStr = nxStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            nx = nx + 1
        Else
            If Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(nxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                nxStr = Replace(nxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(nxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                nxStr = Replace(nxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                nxStr = Replace(nxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "���=" & nxStr
        nxs = nxs + 1
    ElseIf Range(Selection.Address).Value = "��" Then
        findPosi = InStr(Range(posi & posiRow + 5).Value, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            'Range(posi & posiRow + 4 + 1).Value = Range(posi & posiRow + 4 + 1).Value & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Debug.Print Range(posi & posiRow + 5).Value
            zcx = zcx + 1
        End If
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        findPosi = InStr(sjStr, Range("C" & Selection.Row).Value)
        If InStr(sjStr, Range("C" & Selection.Row).Value) <= 0 Then
            sjStr = sjStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            sj = sj + 1
        Else
            If Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(sjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                sjStr = Replace(sjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(sjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                sjStr = Replace(sjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                sjStr = Replace(sjStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "�¼�=" & sjStr
        sjs = sjs + 1
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        findPosi = InStr(bjStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            bjStr = bjStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            bj = bj + 1
        Else
            If Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(bjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bjStr = Replace(bjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(bjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bjStr = Replace(bjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                bjStr = Replace(bjStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & bjStr
        bjs = bjs + 1
    ElseIf Range(Selection.Address).Value = "δǩ��" Then
        findPosi = InStr(qkStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            wqdStr = wqdStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            wqd = wqd + 1
        Else
            If Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(wqdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                wqdStr = Replace(wqdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(wqdStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                wqdStr = Replace(wqdStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(wqdStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                wqdStr = Replace(wqdStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "δǩ��=" & wqdStr
        wqds = wqds + 1
    ElseIf Range(Selection.Address).Value = "��ְ" Then
        If InStr(Range(posi & posiRow + 8 + 1).Value, Range("C" & Selection.Row).Value) <= 0 Then
            Range(posi & posiRow + 8 + 1).Value = Range(posi & posiRow + 8 + 1).Value & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Debug.Print Range(posi & posiRow + 8 + 1).Value
            rz = rz + 1
        Else
        
        End If
    ElseIf Range(Selection.Address).Value = "��ְ" Then
        If InStr(Range(posi & posiRow + 9 + 1).Value, Range("C" & Selection.Row).Value) <= 0 Then
            Range(posi & posiRow + 9 + 1).Value = Range(posi & posiRow + 9 + 1).Value & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Debug.Print Range(posi & posiRow + 9 + 1).Value
            lz = lz + 1
        Else
        
        End If
    ElseIf InStr(Range(Selection.Address).Value, "��") Then '����1 2 3
        findPosi = InStr(dxStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            dxStr = dxStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            dx = dx + 1
        Else
            If Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                ''Debug.Print t1
                ''Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                dxStr = Replace(dxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & dxStr
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
    Selection.Offset(0, 2).Select '����2����Ԫ��
    cellName = Selection.Address
    'Debug.Print "next cell " & cellName
    ColumnName = GetColumnName(Selection.Column)
    'Debug.Print "����--" & Range("C" & Selection.Row).Value & "--�к�--" & ColumnName & Selection.Row & "===" & Range(Selection.Address).Value
    
    If Range(Selection.Address).Value = "����" Then
        findPosi = InStr(ztStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            ztStr = ztStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Range(posi & posiRow + 1).Select
            zt = zt + 1
        Else
            If Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(ztStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                ztStr = Replace(ztStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(ztStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                ztStr = Replace(ztStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(ztStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                ztStr = Replace(ztStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & ztStr
        zts = zts + 1
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        findPosi = InStr(kgStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            kgStr = kgStr & "��" & Range("C" & Selection.Row).Value 'ȡ���� ��һ��
            kg = kg + 1
        Else
            If Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(kgStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                t2 = Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                kgStr = Replace(kgStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(kgStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                t2 = Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                kgStr = Replace(kgStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(kgStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                kgStr = Replace(kgStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & kgStr
        kgs = kgs + 1
    ElseIf Range(Selection.Address).Value = "����" Then
        Selection.Offset(0, -1).Select '���ؿ������Ƿ�Ҳ����
        If Range(Selection.Address).Value <> "����" Then
            findPosi = InStr(dxStr, Range("C" & Selection.Row).Value)
            If findPosi <= 0 Then
                dxStr = dxStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
                Range(posi & posiRow + 4).Select
                dx = dx + 1
            Else
                If Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                    t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                    'Debug.Print t1
                    'Debug.Print Len(Range("C" & Selection.Row).Value)
                    t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                    'Debug.Print t2
                    dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
                ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                    t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                    'Debug.Print t1
                    'Debug.Print Len(Range("C" & Selection.Row).Value)
                    t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                    'Debug.Print t2
                    dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
                ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                    dxStr = Replace(dxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
                End If
            End If
            Debug.Print "����=" & dxStr
            dxs = dxs + 1
        End If
        Selection.Offset(0, 1).Select '�ƻ���
    ElseIf Range(Selection.Address).Value = "��" Then
        Selection.Offset(0, -1).Select '���ؿ������Ƿ�Ҳ����
        If Range(Selection.Address).Value <> "��" Then
            findPosi = InStr(Range(posi & posiRow + 5).Value, Range("C" & Selection.Row).Value)
            If findPosi <= 0 Then
                'Range(posi & posiRow + 4 + 1).Value = Range(posi & posiRow + 4 + 1).Value & "��" & Range("C" & Selection.Row).Value 'ȡ����
                Debug.Print Range(posi & posiRow + 5).Value
                zcx = zcx + 1
            End If
        End If
        Selection.Offset(0, 1).Select '�ƻ���
    ElseIf Range(Selection.Address).Value = "���" Then
        Selection.Offset(0, -1).Select '���ؿ������Ƿ�Ҳ����
        If Range(Selection.Address).Value <> "����" Then
            findPosi = InStr(nxStr, Range("C" & Selection.Row).Value)
            If findPosi <= 0 Then
                nxStr = nxStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
                Range(posi & posiRow + 11).Select
                nx = nx + 1
            Else
                If Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                    t1 = Mid(nxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                    'Debug.Print t1
                    'Debug.Print Len(Range("C" & Selection.Row).Value)
                    t2 = Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                    'Debug.Print t2
                    nxStr = Replace(nxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
                ElseIf Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                    t1 = Mid(nxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                    'Debug.Print t1
                    'Debug.Print Len(Range("C" & Selection.Row).Value)
                    t2 = Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                    'Debug.Print t2
                    nxStr = Replace(nxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
                ElseIf Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(nxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                    nxStr = Replace(nxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
                End If
            End If
            Debug.Print "���=" & nxStr
            nxs = nxs + 1
        End If
        Selection.Offset(0, 1).Select '�ƻ���
    ElseIf Range(Selection.Address).Value = "δǩ��" Then
        findPosi = InStr(wqtStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            wqtStr = wqtStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Range(posi & posiRow + 8).Select
            wqt = wqt + 1
        Else
            If Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(wqtStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                wqtStr = Replace(wqtStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(wqtStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                wqtStr = Replace(wqtStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(wqtStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                wqtStr = Replace(wqtStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "δǩ��=" & wqtStr
        wqts = wqts + 1
    ElseIf Range(Selection.Address).Value = "�Ӱ�" Then '����
        findPosi = InStr(bcStr, Range("C" & Selection.Row).Value)
        'Debug.Print Range("C" & Selection.Row).Value
        'Debug.Print bcStr
        If findPosi <= 0 Then
            'Debug.Print bcStr
            'Debug.Print Range("C" & Selection.Row).Value
            bcStr = bcStr & "��" & Range("C" & Selection.Row).Value 'ȡ���� ��һ��
            bc = bc + 1
        Else
            If Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(bcStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bcStr = Replace(bcStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(bcStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bcStr = Replace(bcStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(bcStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                bcStr = Replace(bcStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & bcStr
        bcs = bcs + 1
    ElseIf Range(Selection.Address).Value = "����" Then
        findPosi = InStr(bqtStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            bqtStr = bqtStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Range(posi & posiRow + 8).Select
            bqt = bqt + 1
        Else
            If Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(bqtStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bqtStr = Replace(bqtStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(bqtStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bqtStr = Replace(bqtStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(bqtStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                bqtStr = Replace(bqtStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "��ǩ��=" & bqtStr
        bqts = bqts + 1
    ElseIf Range(Selection.Address).Value = "���" Then '����
        findPosi = InStr(btStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            btStr = btStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Range(posi & posiRow + 8).Select
            bt = bt + 1
        Else
            If Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(btStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                btStr = Replace(btStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(btStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                btStr = Replace(btStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(btStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                btStr = Replace(btStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & btStr
        bts = bts + 1
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        findPosi = InStr(sjStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            sjStr = sjStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Range(posi & posiRow + 6).Select
            sj = sj + 1
        Else
            If Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(sjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                sjStr = Replace(sjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(sjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                sjStr = Replace(sjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(sjStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                sjStr = Replace(sjStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "�¼�=" & sjStr
        sjs = sjs + 1
    ElseIf InStr(Range(Selection.Address).Value, "��") Then
        findPosi = InStr(bjStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            bjStr = bjStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Range(posi & posiRow + 7).Select
            bj = bj + 1
        Else
            If Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(bjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                bjStr = Replace(bjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(bjStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                bjStr = Replace(bjStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(bjStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                bjStr = Replace(bjStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & bjStr
        bjs = bjs + 1
    ElseIf InStr(Range(Selection.Address).Value, "��") And Len(Range(Selection.Address).Value) > 0 Then
        findPosi = InStr(dxStr, Range("C" & Selection.Row).Value)
        If findPosi <= 0 Then
            dxStr = dxStr & "��" & Range("C" & Selection.Row).Value 'ȡ����
            Range(posi & posiRow + 4).Select
            dx = dx + 1
        Else
            If Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 2, 1) = "��" Then '��λ��
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 2)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 2)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value) + 1, 1) = "��" Then '��λ��
                t1 = Mid(dxStr, findPosi, Len(Range("C" & Selection.Row).Value) + 1)
                'Debug.Print t1
                'Debug.Print Len(Range("C" & Selection.Row).Value)
                t2 = Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1)
                'Debug.Print t2
                dxStr = Replace(dxStr, t1, Range("C" & Selection.Row).Value & t2 + 1)
            ElseIf Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), 1) = "��" Or Mid(dxStr, findPosi + Len(Range("C" & Selection.Row).Value), findPosi + Len(Range("C" & Selection.Row).Value)) = "" Then '�ڶ���
                dxStr = Replace(dxStr, Range("C" & Selection.Row).Value, Range("C" & Selection.Row).Value & "2��")
            End If
        End If
        Debug.Print "����=" & dxStr
        dxs = dxs + 1
    End If
    findPosi = 0
    t1 = ""
    t2 = ""
End Function


Function pdf()
    Dim tempName As String
    'Sheets("����-��ֵ").Select
    'Sheets("2��").Select
    For I = 1 To 18
        tempName = Range("C" & 216 + I).Value
        Range("C202").Value = tempName
        ActiveSheet.Range("$A$6:$BR$201").AutoFilter Field:=3, Criteria1:=tempName
        ChDir "D:\����\�ŶӺ�\202402"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "D:\����\�ŶӺ�\202402\" & I & " " & tempName & "-2024��2�¿���ͳ��.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False
    Next
End Function

Function pdfG()
    Dim tempName As String
    For I = 1 To 318
        tempName = Sheets("����").Range("D" & I + 6).Value
        ActiveSheet.Range("$A$6:$BR$304").AutoFilter Field:=4, Criteria1:=tempName
        ChDir "D:\����\�ŶӺ�\202307-"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "D:\����\�ŶӺ�\202307-\" & I & " " & tempName & "-2022��1�¿���ͳ��.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False
    Next
End Function

Function shiftMark()
    Sheets("����").Select
    Dim ShiftLine(), markColumnName()
    ShiftLine = Array(5, 6, 12, 13, 19, 20, 26, 27, 33, 34)
    markColumnName = Array("K", "M", "Y", "AA", "AM", "AO", "BA", "BC", "BO", "BQ")
    Dim thisName, tempName, nameColumnName As String
    Dim nameColumnStart As Integer
    nameColumnStart = GetColumnNum("AI") '------HK name start column
    For I = 0 To 9 'ֵ������
        'For j = 0 To 16 '������ h to X
        For j = 0 To 3 '������ AI to AL
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
    Sheets("����").Select
    Range(cellName).Select
    If Selection.Interior.Color <> 5287936 Then
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5287936
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Offset(0, 1).Select '����1����Ԫ��
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
            
    Sheets("1��").Select
    totalLine = 176
    
    Dim formatColumnStart, formatColumnEnd, cols As Integer
    Dim tempNum As Double
    formatColumnStart = GetColumnNum("AQ") 'ʵ����ǰһ�� GetColumnNum("BH") '1-15 ���´��ۼ�
    formatColumnEnd = GetColumnNum("BX")
    cols = formatColumnEnd - formatColumnStart
    For Line = 7 To totalLine
        cellName = "AR" & Line
        'Debug.Print (cellName)
        For col = 1 To cols '
            Range(cellName).Select
            Selection.Offset(0, 1).Select '����1����Ԫ��
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
    formatColumnStart = GetColumnNum("AL") 'ʵ����ǰһ�� GetColumnNum("BH") '1-15 ���´��ۼ�
    formatColumnEnd = GetColumnNum("BL")
    cols = formatColumnEnd - formatColumnStart - 2
    For Line = 6 To totalLine
        cellName = "AL" & Line
        'Debug.Print (cellName)
        If Range("F" & Line).Value = "��" Then
            For col = 1 To cols '
                Range(cellName).Select
                If IsNumeric(Selection.Value) Then
                    Selection.Offset(0, 1).Select '����1����Ԫ��
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
    Cells.Replace What:="����1h", Replacement:="��1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����2h", Replacement:="��2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����3h", Replacement:="��3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����4h", Replacement:="��4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����5h", Replacement:="��5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����6h", Replacement:="��6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����7h", Replacement:="��7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��", Replacement:="��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��", Replacement:="��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵδ", Replacement:="δ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��δ", Replacement:="δ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����1h", Replacement:="��1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����2h", Replacement:="��2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����3h", Replacement:="��3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����4h", Replacement:="��4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����5h", Replacement:="��5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����6h", Replacement:="��6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����7h", Replacement:="��7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�¼�1h", Replacement:="��1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�¼�2h", Replacement:="��2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�¼�3h", Replacement:="��3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�¼�4h", Replacement:="��4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�¼�5h", Replacement:="��5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�¼�6h", Replacement:="��6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�¼�7h", Replacement:="��7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����1h", Replacement:="��1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����2h", Replacement:="��2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����3h", Replacement:="��3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����4h", Replacement:="��4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����5h", Replacement:="��5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����6h", Replacement:="��6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����7h", Replacement:="��7", LookAt:=xlPart, _
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
    Cells.Replace What:="����1h", Replacement:="��1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����2h", Replacement:="��2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����3h", Replacement:="��3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����4h", Replacement:="��4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����5h", Replacement:="��5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����6h", Replacement:="��6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����7h", Replacement:="��7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="���1h", Replacement:="��1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="���2h", Replacement:="��2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="���3h", Replacement:="��3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="���4h", Replacement:="��4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="���5h", Replacement:="��5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="���6h", Replacement:="��6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="���7h", Replacement:="��7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��1h", Replacement:="ֵ1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��2h", Replacement:="ֵ2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��3h", Replacement:="ֵ3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��4h", Replacement:="ֵ4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��5h", Replacement:="ֵ5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��6h", Replacement:="ֵ6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ��7h", Replacement:="ֵ7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'    Cells.Replace What:="�ٵ�", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'    Cells.Replace What:="����", Replacement:="-", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
    Cells.Replace What:="m", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="SU(", Replacement:="SUM(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����ֵ��", Replacement:="����", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Function

Function replaceDayHour()
    Cells.Replace What:="1Сʱ", Replacement:="0.125��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="2Сʱ", Replacement:="0.25��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="3Сʱ", Replacement:="0.375��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="4Сʱ", Replacement:="0.5��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="5Сʱ", Replacement:="0.625��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="6Сʱ", Replacement:="0.75��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="7Сʱ", Replacement:="0.875��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Function

Function replaceDay()
    Cells.Replace What:="��1", Replacement:="����0.125��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��2", Replacement:="����0.25��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��3", Replacement:="����0.375��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��4", Replacement:="����0.5��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��5", Replacement:="����0.625��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��6", Replacement:="����0.75��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��7", Replacement:="����0.875��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��1", Replacement:="�¼�0.125��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��2", Replacement:="�¼�0.25��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��3", Replacement:="�¼�0.375��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��4", Replacement:="�¼�0.5��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��5", Replacement:="�¼�0.625��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��6", Replacement:="�¼�0.75��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��7", Replacement:="�¼�0.875��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��1", Replacement:="����0.125��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="������0.125����", Replacement:="����1��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��1", Replacement:="�ٵ�0.125��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��7", Replacement:="����0.875��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����0.125��", Replacement:="�ٵ�0.125��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��4", Replacement:="�ٵ�0.5��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="����0.5��", Replacement:="�ٵ�0.5��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��1", Replacement:="����0.125��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��2", Replacement:="����0.25��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��3", Replacement:="����0.375��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��4", Replacement:="����0.5��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��5", Replacement:="����0.625��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��6", Replacement:="����0.75��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��7", Replacement:="����0.875��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ1", Replacement:="ֵ��0.125��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ2", Replacement:="ֵ��0.25��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ3", Replacement:="ֵ��0.375��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ4", Replacement:="ֵ��0.5��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ5", Replacement:="ֵ��0.625��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ6", Replacement:="ֵ��0.75��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ֵ7", Replacement:="ֵ��0.875��", LookAt:=xlPart, _
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
            If InStr(Sheets("����").Range(monthDayModify(I) & j + 6).Value, "��") > 0 Then
                tmpName = Sheets("����").Range("E" & j + 6).Value
                Debug.Print tmpName
                Sheets("2������").Select
                
                Cells.Find(What:=tmpName, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                    , MatchByte:=False, SearchFormat:=False).Activate
                tmpRow = Selection.Row
                
                Range(monthDay(I) & tmpRow).Select
                Sheets("2������").Range(monthDay(I) & tmpRow).Value = Sheets("����").Range(monthDayModify(I) & j + 6).Value
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
    For I = 1 To 20 '�ݶ�����
        tmpName = Sheets("H").Range("B" & I + 199).Value
        On Error GoTo ErrorHandler
        If tmpName = "" Then
            Exit For
        Else
            Sheets("add").Select
            Columns("G:G").Select
            Selection.Replace What:="����*", Replacement:="--", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            ActiveSheet.Range("$A:$T").AutoFilter Field:=7, Criteria1:=timeStr & " --" '    $A$3:$T$5654
            Cells.Find(What:=tmpName, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                , MatchByte:=False, SearchFormat:=False).Activate
            tmpRow = Selection.Row
            If InStr(Range("K" & tmpRow).Value, "����Դ�Ƽ���չ���޹�˾") > 0 Then
                 While needNext
                    tmpRow = tmpRow + 1
                    If Range("A" & tmpRow).Value = tmpName And Left(Range("G" & tmpRow).Value, 8) = timeStr Then
                        needNext = True
                    Else
                        needNext = False
                        tmpRow = tmpRow - 1
                    End If
                Wend
                If Not InStr(Range("K" & tmpRow).Value, "����Դ�Ƽ���չ���޹�˾") > 0 Then
                    isProject = True
                End If
            Else
                isProject = True
            End If
        End If
        Sheets("H").Select
        If isProject Then
            Range(monthDay(day - 1) & I + 199).Value = "����"
        Else
            Range(monthDay(day - 1) & I + 199).Value = "�칫"
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
    newFileName = "D:\����\�ŶӺ�\����\����1��" & day & "��.xlsx"
    shortName = "����1��" & day & "��.xlsx"
    pdfFileName = "D:\����\�ŶӺ�\����\����1��" & day & "��.pdf"
    Workbooks.Add
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=newFileName, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Windows("1�¿���.xlsm").Activate
    If InStr(weekDay, "��") > 0 Or InStr(weekDay, "��") > 0 Then
        Sheets(Array("d7", "mo")).Select
        Sheets(Array("d7", "mo")).Copy Before:=Workbooks(shortName).Sheets(1)
        Windows(shortName).Activate
        Sheets("sheet1").Delete
        Application.DisplayAlerts = True
        Call clearSheets
        ActiveWorkbook.Save
        Sheets(Array("d7", "mo")).Select
    '    ChDir "D:\����\�ŶӺ�\����"
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
    '    ChDir "D:\����\�ŶӺ�\����"
    '    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    '        pdfFileName, Quality:=xlQualityStandard, IncludeDocProperties _
    '        :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("d").Select
    End If
    ActiveWorkbook.Close
    Windows("1�¿���.xlsm").Activate
    Sheets("1").Select
    Sheets("1��").Select
End Function
Function testOneDay()
    Call oneDay(1, "����")

End Function

Function oneDay(day As Integer, weekDay As String)
    Dim oneDayColumnStart, firstCellStart, secondCellStart, dayNumber As Integer
    Dim firstCell, secondCell, tmpStr As String
    oneDayColumnStart = GetColumnNum("K") '�¶ȱ�1����
    firstCellStart = GetColumnNum("I") '������������������
    secondCellStart = firstCellStart + 1
    dayNumber = day
    Call workDays(dayNumber)
    
    Sheets("1��").Select
    totalLine = 176
    LogNum = 0
    For Line = 7 To totalLine
        Delayms (0.1)
        'If Line = 9 Then callDebug
        For col = dayNumber To dayNumber '
            tmpStr = ""
            Range(GetColumnName(oneDayColumnStart + (col - 1)) & Line).Select
            firstCell = Trim(Sheets("����").Range(GetColumnName(firstCellStart + (col - 1) * 2) & Line).Value)
            secondCell = Trim(Sheets("����").Range(GetColumnName(secondCellStart + (col - 1) * 2) & Line).Value)
            cellName = Selection.Address
            ColumnName = GetColumnName(Selection.Column)
            Debug.Print Line & "��" & Range("D" & Selection.Row).Value & "--�к�--" & ColumnName & Selection.Row & "+++" & Range(Selection.Address).Value
            
            Debug.Print firstCell & "-" & secondCell
            'If (firstCell = "��" And secondCell = "��") Or (firstCell = 0 And secondCell = 0) Then
            If (firstCell = secondCell) Then
                If (firstCell = "��" And secondCell = "��") Then
                    tmpStr = ""
                ElseIf firstCell = "����" And secondCell = "����" Then
                    tmpStr = "��"
                Else
                    tmpStr = firstCell
                End If
            Else
                If secondCell = "��" Then
                    tmpStr = firstCell
                ElseIf firstCell = "��" Then
                    tmpStr = secondCell
                ElseIf IsNumeric(firstCell) And IsNumeric(secondCell) Then
                    tmpStr = firstCell - secondCell
                ElseIf IsNumeric(firstCell) And secondCell = "����" Then
                    tmpStr = firstCell
                ElseIf IsNumeric(firstCell) And secondCell <> "����" Then
                    tmpStr = firstCell & Chr("10") & secondCell
                ElseIf IsNumeric(secondCell) And firstCell = "����" Then
                    tmpStr = secondCell
                ElseIf IsNumeric(secondCell) And firstCell <> "����" Then
                    tmpStr = firstCell & Chr("10") & secondCell
                ElseIf firstCell = "����" And secondCell <> "����" Then
                    If secondCell = "����" Then
                        tmpStr = "��4"
                    ElseIf secondCell = "�¼�" Then
                        tmpStr = "��4"
                    ElseIf secondCell = "����" Then
                        tmpStr = "��4"
                    ElseIf secondCell = "���" Then
                        tmpStr = "��4"
                    ElseIf firstCell = "δǩ��" Then
                        tmpStr = "δǩ��"
                    ElseIf secondCell = "��1" Then
                        tmpStr = "��1"
                    ElseIf secondCell = "��4" Then
                        tmpStr = "��4"
                    Else
                        tmpStr = secondCell
                    End If
                ElseIf firstCell <> "����" And secondCell = "����" Then
                    If firstCell = "����" Then
                        tmpStr = "��4"
                    ElseIf secondCell = "δǩ��" Then
                        tmpStr = "δǩ��"
                    ElseIf firstCell = "�¼�" Then
                        tmpStr = "��4"
                    ElseIf firstCell = "����" Then
                        tmpStr = "��4"
                    ElseIf firstCell = "���" Then
                        tmpStr = "��4"
                    Else
                        tmpStr = firstCell
                    End If
                ElseIf firstCell = "" And secondCell = "����" Then
                    tmpStr = "��4"
                ElseIf firstCell = "����" And secondCell = "" Then
                    tmpStr = "��4"
                ElseIf firstCell = "ֵ��" And secondCell <> "ֵ��" Then
                    If secondCell = "" Or secondCell = 0 Then
                        tmpStr = "ֵ4"
                    Else
                        tmpStr = secondCell
                    End If
                ElseIf firstCell <> "ֵ��" And secondCell = "ֵ��" Then
                    If firstCell = "" Or firstCell = 0 Then
                        tmpStr = "ֵ4"
                    Else
                        tmpStr = firstCell
                    End If
                ElseIf firstCell <> "ֵ�����" And secondCell = "ֵ�����" Then
                    tmpStr = firstCell
                ElseIf firstCell <> "ֵ�����" And secondCell = "ֵ�����" Then
                    tmpStr = firstCell
                Else
                    tmpStr = firstCell & Chr("10") & secondCell
                End If
            End If

            Delayms (0.01)
           ' If Line = 7 Then Range(GetColumnName(oneDayColumnStart + col - 1) & 6) = "1��" & col & "��"
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
            ElseIf InStr(tmpStr, "��1") Or InStr(tmpStr, "��2") Or InStr(tmpStr, "��3") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16751103
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "����") Then
                With Selection.Font
                    .Color = -39169
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "��") Then
                With Selection.Font
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "����") Then
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
            ElseIf InStr(tmpStr, "ɥ��") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark2
                    .TintAndShade = -9.99786370433668E-02
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "ֵ4") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 15773696
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "��1") Or InStr(tmpStr, "��2") Or InStr(tmpStr, "��3") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16764108
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "�¼�") Then
                With Selection.Font
                    .Color = -52327
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "��1") Or InStr(tmpStr, "��2") Or InStr(tmpStr, "��3") Or InStr(tmpStr, "��4") Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 5296274
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "����") Then
                With Selection.Font
                    .Color = -11489280
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "��") Then
                With Selection.Font
                    .Color = -65536
                    .TintAndShade = 0
                End With
            ElseIf InStr(tmpStr, "��") Then
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
            ElseIf InStr(tmpStr, "δ") Then
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


'1 ?ͨ��������ת���ɶ�Ӧ���к�?
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
'����:

Function TestGetColumnNum()
    Dim ColumnNum As Integer
     ColumnNum = GetColumnNum("ET")
     Debug.Print ColumnNum
     'MsgBox ColumnNum, vbInformation, "����"
 End Function
 
'2 ?ͨ���к�ת���ɶ�Ӧ��������?
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
'����:
'Function TestGetColumnName()
'     Dim ColumnName As String
'     ColumnName = GetColumnName(54)
'    MsgBox ColumnName, vbInformation, "����"
' End Function
'3 ?˵��
'���������������У��������Ĳ�������Excel������к�"IV"(256),�򷵻ص�ֵΪ����������
'��������������������������������
'��Ȩ����������ΪCSDN������xuanxingmin����ԭ�����£���ѭCC 4.0 BY-SA��ȨЭ�飬ת���븽��ԭ�ĳ������Ӽ���������
'ԭ�����ӣ�https://blog.csdn.net/xuanxingmin/article/details/2582861

Function statusTest()
    Sheets("day").Select
    Dim coName(), StatusName() As Variant, temp, rangePosi As String, flagRow, flagCol As Integer
    
    coName() = Range("A6:A23").Value
    StatusName() = Range("I5:BB5").Value
    
    temp = "����"
    
    On Error Resume Next
    flagRow = WorksheetFunction.Match(temp, coName, 0) + 5 '������5��
    Debug.Print flagRow
        
    temp = "�Ӱ�"
    On Error Resume Next
    flagCol = WorksheetFunction.Match(temp, StatusName, 0) + 8 'ǰ����8��
    rangePosi = GetColumnName(flagCol) & flagRow
    Debug.Print rangePosi
    
    Debug.Print Range(rangePosi).Value
    
End Function

Function IsWeekNormal(ByVal searchValue As String) As Boolean
    Sheets("����").Select
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
    Debug.Print removeNonDigits1("gһ1.207")
End Function

Function KeepNumbersAndDecimals(strInput As String) As String
    ' ʹ��������ʽ�滻�����ֺͷ�С������ַ�Ϊ��
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .Pattern = "[^\d.]"
        KeepNumbersAndDecimals = .Replace(strInput, vbNullString)
    End With
End Function

Function callDebug()
    Application.EnableCancelKey = xlInterrupt '����ȡ����Ϊ�ж�״̬
    Debug.Assert False '��������Ϊ�ϵ��������Ҫ���Ե�λ����
End Function

Function SortStr(str As String) As String
    Dim arr As Variant

    '����Ҫ���������
'    arr = Array(5, 2, 8, 1, 9)
    arr = Split(str, Chr(10))

    '���� Sort �����������������Ĭ��Ϊ����
    Call Sort(arr)

    '��������Ľ��
    For I = LBound(arr) To UBound(arr)
        Debug.Print arr(I)
        If I = LBound(arr) Then
            SortStr = arr(I)
        Else
            SortStr = SortStr & Chr(10) & arr(I)
        End If
    Next I
End Function
 
'�Զ���� Sort ����
Function Sort(ByRef arr As Variant)
    Dim tempArr As Variant
    Dim I As Integer, j As Integer
    
    '����������鸳ֵ����ʱ����
    tempArr = arr
    
    'ͨ���Ƚ�����Ԫ�ز�����λ�����������
    For I = LBound(tempArr) To UBound(tempArr) - 1
        For j = I + 1 To UBound(tempArr)
            If tempArr(j) < tempArr(I) Then
                SwapElements tempArr, I, j
            End If
        Next j
    Next I
    
    '������õ��������¸�ֵ��ԭʼ����
    arr = tempArr
End Function
 
'��������Ԫ�ص�λ��
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
    If n > 0 Then MsgBox "�ı�������ת��Ϊ��ֵ�ɹ���", vbOKOnly, "�ɹ�" Else MsgBox "δ��⵽�ı������֣�", vbOKOnly + vbCritical, "����"

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
    If n > 0 Then MsgBox "�ı�������ת��Ϊ��ֵ�ɹ���", vbOKOnly, "�ɹ�" Else MsgBox "δ��⵽�ı������֣�", vbOKOnly + vbCritical, "����"
    Set rng = Nothing
End Sub
