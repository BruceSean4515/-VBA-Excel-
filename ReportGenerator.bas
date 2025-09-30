Attribute VB_Name = "ģ��1"
' =============================================================================
' ��̽��Ŀ�ձ��Զ����ɹ��ߣ�ͨ��ģ�壩
' ���ܣ���Excel��ȡ��̽���ݣ��Զ����������ձ�������ΪWord�ĵ�
' ˵����������Ϊͨ��ģ�壬�����κ���ʵ��Ŀ��Ϣ����ֱ������ѧϰ����ο���
' ���ߣ�[BruceSean4515]
' ���֤��MIT License
' =============================================================================

Sub ����ÿ�հ౨������Word()
    On Error GoTo ErrorHandler
    
    Dim wsData As Worksheet
    Dim reportText As String
    Dim rowMachine As Range
    Dim machineNo As String, machineType As String, holeNo As Variant, designDepth As Variant, currentDepth As Variant, dailyFootage As Variant, startDate As Variant
    Dim completedHoles As Variant, completedDepth As Variant, remarks As String
    Dim todayCompletedHoles As Variant, todayFootageTotal As Variant
    Dim drillCount As Variant, personCount As Variant, holeCount As Variant, totalDepth As Variant, workingHolesText As String, inHoleDepth As Variant, todayFootage As Variant
    Dim reportDate As String, weather As String
    Dim savePath As String, fileName As String, fullFilePath As String
    Dim wdApp As Object, wdDoc As Object
    
    ' �������ݹ�������ȷ������������Ϊ���ճ����¡���
    Set wsData = ThisWorkbook.Sheets("�ճ�����")
    
    ' ���������򡰻���ͳ���С���ȡ��Ŀ��������
    With wsData
        reportDate = Format(.Range("����ͳ����").Cells(1, 1).Value, "m��d��") ' A�У���������
        weather = Nz(.Range("����ͳ����").Cells(1, 2).Value, "��")           ' B�У�����
        drillCount = .Range("����ͳ����").Cells(1, 3).Value                  ' C�У��������
        personCount = .Range("����ͳ����").Cells(1, 4).Value                 ' D�У��ֳ���Ա��
        holeCount = .Range("����ͳ����").Cells(1, 5).Value                   ' E�У��ۼ���������
        totalDepth = .Range("����ͳ����").Cells(1, 6).Value                  ' F�У��ۼ���̽���ߣ��ף�
        workingHolesText = .Range("����ͳ����").Cells(1, 7).Value            ' G�У�����ʩ���������
        inHoleDepth = .Range("����ͳ����").Cells(1, 8).Value                 ' H�У������ܽ��ߣ��ף�
        todayFootage = .Range("����ͳ����").Cells(1, 9).Value                ' I�У������ܽ��ߣ��ף�
    End With
    
    ' �ӡ�����ʩ����ס��ı�����ȡ���֣����硰3���� �� 3��
    Dim workingHoles As Long
    workingHoles = 0
    If Not IsError(workingHolesText) Then
        workingHoles = Val(workingHolesText)
    End If
    
    ' �����ձ�����ͻ��ܶ��䣨ʹ��ͨ����Ŀ���ƣ�
    reportText = "xx��̽��Ŀ�ձ���" & reportDate & "��" & vbCrLf
    reportText = reportText & weather & "�����" & Nz(drillCount, 0) & "̨����Ա:" & Nz(personCount, 0) & "����" & _
                 "�ۼ�������" & Nz(holeCount, 0) & "���������̽������:" & FormatNumber(Nz(totalDepth, 0), 2) & "m��" & _
                 "����ʩ�����:" & workingHoles & "�������ڽ���:" & FormatNumber(Nz(inHoleDepth, 0), 2) & "m��" & _
                 "���ս���" & FormatNumber(Nz(todayFootage, 0), 2) & "m��" & vbCrLf & vbCrLf
    
    ' ��������̨���ݱ��е�ÿһ�л�̨����
    For Each rowMachine In wsData.Range("��̨���ݱ�").Rows
        machineNo = Nz(rowMachine.Cells(1, 1).Value, "") ' A�У���̨���
        If machineNo <> "" And Not IsError(machineNo) Then
            
            remarks = Nz(rowMachine.Cells(1, 11).Value, "") ' K�У���̨״̬�����ڳ���/���ѳ�����/�գ�
            
            ' ����δ���õĻ�̨����עΪ�գ�
            If remarks = "" Then GoTo NextMachine
            
            machineType = Nz(rowMachine.Cells(1, 2).Value, "��δ��д��") ' B�У��������
            completedHoles = Nz(rowMachine.Cells(1, 3).Value, 0) ' C�У��ۼ���������
            completedDepth = Nz(rowMachine.Cells(1, 4).Value, 0) ' D�У��ۼ��տ׽��ߣ��ף�
            holeNo = Nz(rowMachine.Cells(1, 5).Value, "") ' E�У���ǰʩ����ױ��
            designDepth = Nz(rowMachine.Cells(1, 6).Value, 0) ' F�У���ƿ���ף�
            currentDepth = Nz(rowMachine.Cells(1, 7).Value, 0) ' G�У���ǰ����ף�
            dailyFootage = Nz(rowMachine.Cells(1, 8).Value, 0) ' H�У����ս��ߣ��ף�
            startDate = Nz(rowMachine.Cells(1, 9).Value, "") ' I�У���������
            
            ' ��ȡ�����տ�������ݣ�O�С�P�У�
            todayCompletedHoles = Nz(rowMachine.Cells(1, 15).Value, 0) ' O�У������տ�����
            todayFootageTotal = Nz(rowMachine.Cells(1, 16).Value, 0) ' P�У��û�̨�����ܽ��ߣ��ף�
            
            ' д���̨������Ϣ
            reportText = reportText & machineNo & "�Ż���" & machineType & "����" & _
                         "�ۼ�������:" & completedHoles & "������̽������" & FormatNumber(completedDepth, 2) & "m��"
            
            ' �ѳ������
            If remarks = "�ѳ���" Then
                reportText = reportText & "��ɸ���Ŀ�����ѳ�����" & vbCrLf & vbCrLf
                GoTo NextMachine
            End If
            
            ' �ڳ����
            If remarks = "�ڳ�" Then
                ' �����տ���Ϣ
                If todayCompletedHoles > 0 Then
                    If todayCompletedHoles = 1 Then
                        reportText = reportText & "�����տ�1�����ÿ׵��ս���" & FormatNumber(todayFootageTotal, 2) & "m��"
                    Else
                        reportText = reportText & "�����տ�" & todayCompletedHoles & "�������ս���" & FormatNumber(todayFootageTotal, 2) & "m��"
                    End If
                End If
                
                ' ��ǰʩ�������Ϣ
                If holeNo <> "" Then
                    reportText = reportText & "��ʩ�����" & holeNo & "��" & _
                                 "��ƿ���" & FormatNumber(designDepth, 2) & "m��" & _
                                 "����" & FormatNumber(currentDepth, 2) & "m��" & _
                                 "���ս���" & FormatNumber(dailyFootage, 2) & "m��"
                    
                    If IsDate(startDate) Then
                        reportText = reportText & Format(startDate, "m��d��") & "���ף����������" & vbCrLf & vbCrLf
                    Else
                        reportText = reportText & "��������δ����������" & vbCrLf & vbCrLf
                    End If
                Else
                    reportText = reportText & vbCrLf & vbCrLf
                End If
            End If
        End If
        
NextMachine:
    Next rowMachine
    
    ' ����ձ���β��
    reportText = reportText & "��̽�豸������������Ա��פ�ذ�ȫ��"
    
    ' === �ļ�����·����ͨ�û�����===
    ' Ĭ�ϱ��浽�û�����ġ���̽�ձ����ļ��У�����Ӳ����·��
    savePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\��̽�ձ�\"
    
    ' �Զ������ļ��У���������ڣ�
    If Dir(savePath, vbDirectory) = "" Then
        MkDir savePath
    End If
    
    ' �����ļ�������ʽ����̽�ձ�_20250405.docx��
    fileName = "��̽�ձ�_" & Format(wsData.Range("����ͳ����").Cells(1, 1).Value, "yyyymmdd") & ".docx"
    fullFilePath = savePath & fileName
    
    ' ���� Word Ӧ�ó���
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Add
    
    ' д���ı������ø�ʽ
    With wdDoc
        .Content.Text = reportText
        .Content.Font.Name = "΢���ź�"
        .Content.Font.Size = 12
        
        ' ����һ����Ϊ���⣨�Ӵ�+�ֺ�14��
        If .Paragraphs.Count > 0 Then
            .Paragraphs(1).Range.Font.Bold = True
            .Paragraphs(1).Range.Font.Size = 14
        End If
        
        .SaveAs2 fullFilePath
        .Close
    End With
    
    ' �������
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    MsgBox "�ձ��ѳɹ����ɣ�" & vbCrLf & "�ļ�����λ�ã�" & fullFilePath, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "�������г���" & vbCrLf & Err.Description & vbCrLf & _
           "���飺1. �����������Ƿ�Ϊ���ճ����¡���2. �Ƿ��Ѷ����������򡰻���ͳ���С��͡���̨���ݱ���", vbCritical
End Sub

' ������������ȫ�����ֵ������ֵ
Function Nz(v As Variant, Optional defaultValue As Variant = 0) As Variant
    If IsError(v) Then
        Nz = defaultValue
    ElseIf v = "" Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

' ������������ȫ��ʽ�����֣�����ָ��С��λ��
Function FormatNumber(v As Variant, decimals As Integer) As String
    Dim fmt As String
    fmt = "0." & String(decimals, "0")
    If IsNumeric(v) And Not IsError(v) Then
        FormatNumber = Format(v, fmt)
    Else
        FormatNumber = Format(0, fmt)
    End If
End Function

