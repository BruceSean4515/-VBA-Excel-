Attribute VB_Name = "模块1"
' =============================================================================
' 钻探项目日报自动生成工具（通用模板）
' 功能：从Excel读取钻探数据，自动生成中文日报并导出为Word文档
' 说明：本代码为通用模板，不含任何真实项目信息，可直接用于学习或二次开发
' 作者：[BruceSean4515]
' 许可证：MIT License
' =============================================================================

Sub 生成每日班报并导出Word()
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
    
    ' 设置数据工作表（请确保工作表名称为“日常更新”）
    Set wsData = ThisWorkbook.Sheets("日常更新")
    
    ' 从命名区域“汇总统计行”读取项目汇总数据
    With wsData
        reportDate = Format(.Range("汇总统计行").Cells(1, 1).Value, "m月d日") ' A列：报告日期
        weather = Nz(.Range("汇总统计行").Cells(1, 2).Value, "晴")           ' B列：天气
        drillCount = .Range("汇总统计行").Cells(1, 3).Value                  ' C列：钻机数量
        personCount = .Range("汇总统计行").Cells(1, 4).Value                 ' D列：现场人员数
        holeCount = .Range("汇总统计行").Cells(1, 5).Value                   ' E列：累计完成钻孔数
        totalDepth = .Range("汇总统计行").Cells(1, 6).Value                  ' F列：累计钻探进尺（米）
        workingHolesText = .Range("汇总统计行").Cells(1, 7).Value            ' G列：正在施工钻孔描述
        inHoleDepth = .Range("汇总统计行").Cells(1, 8).Value                 ' H列：孔内总进尺（米）
        todayFootage = .Range("汇总统计行").Cells(1, 9).Value                ' I列：当日总进尺（米）
    End With
    
    ' 从“正在施工钻孔”文本中提取数字（例如“3个” → 3）
    Dim workingHoles As Long
    workingHoles = 0
    If Not IsError(workingHolesText) Then
        workingHoles = Val(workingHolesText)
    End If
    
    ' 构建日报标题和汇总段落（使用通用项目名称）
    reportText = "xx钻探项目日报（" & reportDate & "）" & vbCrLf
    reportText = reportText & weather & "，钻机" & Nz(drillCount, 0) & "台，人员:" & Nz(personCount, 0) & "名，" & _
                 "累计完成钻孔" & Nz(holeCount, 0) & "个，完成钻探工作量:" & FormatNumber(Nz(totalDepth, 0), 2) & "m。" & _
                 "正在施工钻孔:" & workingHoles & "个，孔内进尺:" & FormatNumber(Nz(inHoleDepth, 0), 2) & "m，" & _
                 "当日进尺" & FormatNumber(Nz(todayFootage, 0), 2) & "m。" & vbCrLf & vbCrLf
    
    ' 遍历“机台数据表”中的每一行机台数据
    For Each rowMachine In wsData.Range("机台数据表").Rows
        machineNo = Nz(rowMachine.Cells(1, 1).Value, "") ' A列：机台编号
        If machineNo <> "" And Not IsError(machineNo) Then
            
            remarks = Nz(rowMachine.Cells(1, 11).Value, "") ' K列：机台状态（“在场”/“已撤场”/空）
            
            ' 跳过未启用的机台（备注为空）
            If remarks = "" Then GoTo NextMachine
            
            machineType = Nz(rowMachine.Cells(1, 2).Value, "（未填写）") ' B列：钻机类型
            completedHoles = Nz(rowMachine.Cells(1, 3).Value, 0) ' C列：累计完成钻孔数
            completedDepth = Nz(rowMachine.Cells(1, 4).Value, 0) ' D列：累计终孔进尺（米）
            holeNo = Nz(rowMachine.Cells(1, 5).Value, "") ' E列：当前施工钻孔编号
            designDepth = Nz(rowMachine.Cells(1, 6).Value, 0) ' F列：设计孔深（米）
            currentDepth = Nz(rowMachine.Cells(1, 7).Value, 0) ' G列：当前孔深（米）
            dailyFootage = Nz(rowMachine.Cells(1, 8).Value, 0) ' H列：当日进尺（米）
            startDate = Nz(rowMachine.Cells(1, 9).Value, "") ' I列：开孔日期
            
            ' 读取当日终孔相关数据（O列、P列）
            todayCompletedHoles = Nz(rowMachine.Cells(1, 15).Value, 0) ' O列：当日终孔数量
            todayFootageTotal = Nz(rowMachine.Cells(1, 16).Value, 0) ' P列：该机台当日总进尺（米）
            
            ' 写入机台基础信息
            reportText = reportText & machineNo & "号机（" & machineType & "），" & _
                         "累计完成钻孔:" & completedHoles & "个，钻探工作量" & FormatNumber(completedDepth, 2) & "m。"
            
            ' 已撤场情况
            If remarks = "已撤场" Then
                reportText = reportText & "完成该项目任务，已撤场。" & vbCrLf & vbCrLf
                GoTo NextMachine
            End If
            
            ' 在场情况
            If remarks = "在场" Then
                ' 当日终孔信息
                If todayCompletedHoles > 0 Then
                    If todayCompletedHoles = 1 Then
                        reportText = reportText & "今日终孔1个，该孔当日进尺" & FormatNumber(todayFootageTotal, 2) & "m。"
                    Else
                        reportText = reportText & "今日终孔" & todayCompletedHoles & "个，当日进尺" & FormatNumber(todayFootageTotal, 2) & "m。"
                    End If
                End If
                
                ' 当前施工钻孔信息
                If holeNo <> "" Then
                    reportText = reportText & "现施工钻孔" & holeNo & "，" & _
                                 "设计孔深" & FormatNumber(designDepth, 2) & "m，" & _
                                 "孔深" & FormatNumber(currentDepth, 2) & "m，" & _
                                 "当日进尺" & FormatNumber(dailyFootage, 2) & "m，"
                    
                    If IsDate(startDate) Then
                        reportText = reportText & Format(startDate, "m月d日") & "开孔，正常钻进。" & vbCrLf & vbCrLf
                    Else
                        reportText = reportText & "开孔日期未填，正常钻进。" & vbCrLf & vbCrLf
                    End If
                Else
                    reportText = reportText & vbCrLf & vbCrLf
                End If
            End If
        End If
        
NextMachine:
    Next rowMachine
    
    ' 添加日报结尾语
    reportText = reportText & "钻探设备运行正常，人员及驻地安全。"
    
    ' === 文件保存路径（通用化处理）===
    ' 默认保存到用户桌面的“钻探日报”文件夹，避免硬编码路径
    savePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\钻探日报\"
    
    ' 自动创建文件夹（如果不存在）
    If Dir(savePath, vbDirectory) = "" Then
        MkDir savePath
    End If
    
    ' 生成文件名（格式：钻探日报_20250405.docx）
    fileName = "钻探日报_" & Format(wsData.Range("汇总统计行").Cells(1, 1).Value, "yyyymmdd") & ".docx"
    fullFilePath = savePath & fileName
    
    ' 启动 Word 应用程序
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Add
    
    ' 写入文本并设置格式
    With wdDoc
        .Content.Text = reportText
        .Content.Font.Name = "微软雅黑"
        .Content.Font.Size = 12
        
        ' 将第一段设为标题（加粗+字号14）
        If .Paragraphs.Count > 0 Then
            .Paragraphs(1).Range.Font.Bold = True
            .Paragraphs(1).Range.Font.Size = 14
        End If
        
        .SaveAs2 fullFilePath
        .Close
    End With
    
    ' 清理对象
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    MsgBox "日报已成功生成！" & vbCrLf & "文件保存位置：" & fullFilePath, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "程序运行出错：" & vbCrLf & Err.Description & vbCrLf & _
           "请检查：1. 工作表名称是否为“日常更新”；2. 是否已定义命名区域“汇总统计行”和“机台数据表”。", vbCritical
End Sub

' 辅助函数：安全处理空值、错误值
Function Nz(v As Variant, Optional defaultValue As Variant = 0) As Variant
    If IsError(v) Then
        Nz = defaultValue
    ElseIf v = "" Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

' 辅助函数：安全格式化数字（保留指定小数位）
Function FormatNumber(v As Variant, decimals As Integer) As String
    Dim fmt As String
    fmt = "0." & String(decimals, "0")
    If IsNumeric(v) And Not IsError(v) Then
        FormatNumber = Format(v, fmt)
    Else
        FormatNumber = Format(0, fmt)
    End If
End Function

