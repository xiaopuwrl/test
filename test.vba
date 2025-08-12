Function GenerateSchedule(employees As Variant, weeks As Variant) As Variant
    ' 初始化结果数组
    Dim schedule() As Variant
    ReDim schedule(1 To UBound(weeks), 1 To 4) ' 每周4个项目
    
    Dim currentGroup As Integer
    Dim currentPosition(1 To 4) As Integer ' 跟踪每组当前的位置
    
    Dim groupCnt As Integer
    groupCnt = (UBound(employees) - LBound(employees)) / 4
    
    ' 初始化当前组和位置
    currentGroup = 1
    For i = 1 To 4
        currentPosition(i) = 1
    Next i
    
    For weekIndex = 1 To UBound(weeks)
        Dim weekNum As Integer
        weekNum = weeks(weekIndex)
        
        ' 每4周切换组
        If weekNum > 1 And (weekNum - 1) Mod 4 = 0 Then
            currentGroup = currentGroup + 1
            currentGroup = (currentGroup - 1) Mod groupCnt + 1
            For i = 1 To 4
                currentPosition(i) = 1
            Next i
        End If
        
        ' 确定当前组的起始索引
        Dim groupStart As Integer
        groupStart = 1 + (currentGroup - 1) * 4
        
        ' 为每个项目分配人员
        For projectNum = 1 To 4 ' A=1, B=2, C=3, D=4
            Dim projectChar As String
            projectChar = Chr(64 + projectNum) ' 转换为A,B,C,D
            
            Dim found As Boolean
            found = False
            Dim attempts As Integer
            attempts = 0
            
            ' 尝试找到合适的人员
            Do While Not found And attempts < 4
                Dim candidateIndex As Integer
                candidateIndex = groupStart - 1 + currentPosition(projectNum)
                
                ' 检查候选人是否适合该项目
                If IsEmployeeSuitable(employees(candidateIndex), projectChar) Then
                    schedule(weekIndex, projectNum) = candidateIndex
                    found = True
                    ' 移动到下一个位置
                    currentPosition(projectNum) = currentPosition(projectNum) + 1
                    If currentPosition(projectNum) > 4 Then currentPosition(projectNum) = 1
                Else
                    ' 尝试下一个候选人
                    currentPosition(projectNum) = currentPosition(projectNum) + 1
                    If currentPosition(projectNum) > 4 Then currentPosition(projectNum) = 1
                    attempts = attempts + 1
                End If
            Loop
            
            ' 如果没有找到合适的人，选择第一个可以的人（应该不会发生，如果输入合理）
            If Not found Then
                For i = 1 To 4
                    candidateIndex = groupStart - 1 + i
                    If IsEmployeeSuitable(employees(candidateIndex), projectChar) Then
                        schedule(weekIndex, projectNum) = candidateIndex
                        currentPosition(projectNum) = i + 1
                        If currentPosition(projectNum) > 4 Then currentPosition(projectNum) = 1
                        Exit For
                    End If
                Next i
            End If
        Next projectNum
        
        ' 如果同一人被分配到多个项目，需要调整（简化处理）
        AdjustConflicts schedule, weekIndex, employees, groupStart
    Next weekIndex
    
    GenerateSchedule = schedule
End Function

Function IsEmployeeSuitable(employeeSkills As Variant, project As String) As Boolean
    ' 检查员工是否适合该项目
    If IsEmpty(employeeSkills) Then
        IsEmployeeSuitable = True ' 如果没有限制，可以任何项目
        Exit Function
    End If
    
    For Each skill In employeeSkills
        If skill = project Then
            IsEmployeeSuitable = True
            Exit Function
        End If
    Next skill
    
    IsEmployeeSuitable = False
End Function

Sub AdjustConflicts(schedule() As Variant, ByVal weekIndex As Integer, employees As Variant, groupStart As Integer)
    ' 检查并解决同一人被分配到多个项目的冲突
    Dim usedEmployees(1 To 4) As Boolean
    Dim conflicts As Boolean
    conflicts = True
    
    ' 尝试解决冲突（最多尝试4次）
    Dim attempts As Integer
    attempts = 0
    
    Do While conflicts And attempts < 10
        conflicts = False
        Erase usedEmployees
        
        For projectNum = 1 To 4
            Dim empIndex As Integer
            empIndex = schedule(weekIndex, projectNum) - (groupStart - 1)
            
            If empIndex >= 1 And empIndex <= 4 Then
                If usedEmployees(empIndex) Then
                    ' 冲突发生，尝试为该项目找到下一个合适的人
                    conflicts = True
                    Dim originalPos As Integer
                    originalPos = empIndex
                    
                    Dim foundAlternative As Boolean
                    foundAlternative = False
                    Dim tryCount As Integer
                    tryCount = 0
                    
                    Do While Not foundAlternative And tryCount < 10
                        originalPos = originalPos + 1
                        If originalPos > 4 Then originalPos = 1
                        
                        If Not usedEmployees(originalPos) Then
                            Dim projectChar As String
                            projectChar = Chr(64 + projectNum)
                            
                            If IsEmployeeSuitable(employees(groupStart - 1 + originalPos), projectChar) Then
                                schedule(weekIndex, projectNum) = groupStart - 1 + originalPos
                                foundAlternative = True
                            End If
                        End If
                        
                        tryCount = tryCount + 1
                    Loop
                End If
                
                If Not conflicts Then
                    usedEmployees(empIndex) = True
                End If
            End If
        Next projectNum
        
        attempts = attempts + 1
    Loop
End Sub

Sub GenerateRoster()
    ' 从Excel读取员工数据 (假设在Sheet1的A1:C4范围)
    Dim employeeRange As Range
    Set employeeRange = ThisWorkbook.Worksheets("Sheet1").Range("A1:C4")
    
    ' 周数数组 (1-52周)
    Dim weeks(1 To 52) As Integer
    For i = 1 To 52
        weeks(i) = i
    Next i
    
    ' 生成排班计划
    Dim schedule As Variant
    schedule = GenerateSchedule(employeeRange, weeks)
    
    ' 输出结果到新工作表
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = "排班结果"
    
    ' 写标题
    ws.Cells(1, 1).Value = "周数"
    ws.Cells(1, 2).Value = "项目A"
    ws.Cells(1, 3).Value = "项目B"
    ws.Cells(1, 4).Value = "项目C"
    ws.Cells(1, 5).Value = "项目D"
    
    ' 写数据
    For week = 1 To UBound(weeks)
        ws.Cells(week + 1, 1).Value = "第" & week & "周"
        For project = 1 To 4
            ws.Cells(week + 1, project + 1).Value = schedule(week, project)
        Next project
    Next week
    
    ' 自动调整列宽
    ws.Columns.AutoFit
    
    MsgBox "排班计划生成完成！", vbInformation
End Sub

Sub TestScheduleGenerator()
    ' 示例员工数组 (8个员工，前4个是第一组，后4个是第二组)
    Dim employees(1 To 16) As Variant
    ' 员工1: 可以值班A,B,C
    employees(1) = Array("A", "B", "C")
    ' 员工2: 可以值班A,B,C,D (无限制)
    employees(2) = Empty
    ' 员工3: 可以值班B,C,D
    employees(3) = Array("B", "C", "D")
    ' 员工4: 只能值班D
    employees(4) = Array("D")
    ' 第二组员工...
    employees(5) = Array("A", "B", "C", "D")
    employees(6) = Array("A", "B", "C", "D")
    employees(7) = Array("A", "B", "C", "D")
    employees(8) = Array("A", "B", "C", "D")
        employees(9) = Array("A", "B", "C", "D")
    employees(10) = Array("A", "B", "C", "D")
    employees(11) = Array("A", "B", "C", "D")
    employees(12) = Array("A", "B", "C", "D")
        employees(13) = Array("A", "B", "C", "D")
    employees(14) = Array("A", "B", "C", "D")
    employees(15) = Array("A", "B", "C", "D")
    employees(16) = Array("A", "B", "C", "D")
    
    ' 周数数组 (1-52周)
    Dim weeks(1 To 52) As Integer
    For i = 1 To 52
        weeks(i) = i
    Next i
    
    ' 生成排班计划
    Dim schedule As Variant
    schedule = GenerateSchedule(employees, weeks)
    
    ' 输出结果 (示例: 输出前4周)
    For week = 1 To 24
        Debug.Print "Week " & week & ": " & _
                    "A:" & schedule(week, 1) & ", " & _
                    "B:" & schedule(week, 2) & ", " & _
                    "C:" & schedule(week, 3) & ", " & _
                    "D:" & schedule(week, 4)
    Next week
End Sub
