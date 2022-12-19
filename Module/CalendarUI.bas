Attribute VB_Name = "CalendarUI"
' Macro Apply : [FileName]!'Launch_CalrndarUI (Left, Top, Height, Width)'
' If "Top","Left" < 0 then, positioning value center of "Height","Width".
Private objTargetDate As Variant

Private Type colorElementType
    LightColor As Long
    DarkColor As Long
End Type

Private defaultColor As colorElementType
Private accentColor As colorElementType
Private shadowColor As Long

Private Const DateFormat As String = "YYYY/M/D"


Private Function Init_CalendarUI()
    defaultColor.LightColor = RGB(255, 255, 255)
    defaultColor.DarkColor = RGB(0, 0, 0)
    
    accentColor.LightColor = RGB(252, 252, 252)
    accentColor.DarkColor = RGB(3, 89, 150)
    
    shadowColor = RGB(127, 127, 127)
    
    ' ==== Date I/O Target Object =====
        ' ---- Cells ----
        Set objTargetDate = ActiveSheet.Cells(2, 10)

        ' ---- Shape ----
        'Set objTargetDate = ActiveSheet.Shapes("shape_00").TextFrame.Characters
        
End Function


Public Sub Lanuch_CalendarUI( _
            Optional Left As Single = 10, _
            Optional Top As Single = 10, _
            Optional Height As Single = 200, _
            Optional Width As Single = 200 _
            )
                                                       
'    Application.ScreenUpdating = False
'    Application.Cursor = xlWait
'    Application.EnableEvents = False
'    Application.DisplayAlerts = False
'    Application.Calculation = xlCalculationManual
    
    Call Init_CalendarUI
    
    Left = Position_Centering(Left, Width)
    Top = Position_Centering(Top, Height)
    
    Call Calendar_Show(Left, Top, Height, Width)

    
'    Application.Calculation = xlCalculationAutomatic
'    Application.DisplayAlerts = True
'    Application.EnableEvents = True
'    Application.Cursor = xlDefault
'    Application.ScreenUpdating = True
    
End Sub


Public Function Calendar_DayClick(targetDayPosition As Integer)
    Dim targetShape As Shape
    Dim i As Integer
    
    For i = 1 To 42
        Set targetShape = ActiveSheet.Shapes("Calendar_Button_Day_" + Format(i, "00"))
        
        With targetShape
            If .TextFrame.Characters.Font.Color <> shadowColor Then
                If i = targetDayPosition Then
                    .OnAction = "'Calendar_DayClick_2Times " + CStr(i) + "'"
                    .Fill.ForeColor.RGB = accentColor.DarkColor
                    .TextFrame.Characters.Font.Color = accentColor.LightColor
                Else
                    .OnAction = "'Calendar_DayClick " + CStr(i) + "'"
                    .Fill.ForeColor.RGB = defaultColor.LightColor
                    .TextFrame.Characters.Font.Color = defaultColor.DarkColor
                End If
            End If
        End With
    Next i
        
End Function


Public Function Calendar_DayClick_2Times(targetDayPosition As Integer)
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim targetDay As Integer
    
    With ActiveSheet
        With .Shapes("Calendar_TextBox_Date")
            targetYear = CInt(Mid(.TextFrame.Characters.Text, 1, 4))
            targetMonth = CInt(Mid(.TextFrame.Characters.Text, 8))
        End With
        
        targetDay = .Shapes("Calendar_Button_Day_" + Format(targetDayPosition, "00")).TextFrame.Characters.Text
    
    End With
    
    Call Calendar_Delete
    
    On Error GoTo TypeRange
        objTargetDate.Text = Format(DateSerial(targetYear, targetMonth, targetDay), DateFormat)
    Exit Function

TypeRange: ' Range don't have method "Text"
    objTargetDate.value = Format(DateSerial(targetYear, targetMonth, targetDay), DateFormat)
    
    
End Function


Public Function Calendar_MonthClick(diffValue As Integer)
    Dim buffDate As Date
    Dim buffDateYear As Integer
    Dim buffDateMonth As Integer
    
    With ActiveSheet.Shapes("Calendar_TextBox_Date")
        buffDateYear = CInt(Mid(.TextFrame.Characters.Text, 1, 4))
        buffDateMonth = CInt(Mid(.TextFrame.Characters.Text, 8))
    End With
    
    buffDate = DateAdd("m", diffValue, DateSerial(buffDateYear, buffDateMonth, 1))
    Call Calendar_DayDraw(buffDate)
    
    With ActiveSheet.Shapes("Calendar_TextBox_Date")
        .TextFrame.Characters.Text = CStr(Year(buffDate)) + " . " + CStr(Month(buffDate))
    End With

End Function


Private Function Position_Centering(targetValue As Single, referenceValue As Single) As Single
    Dim rtnSingle As Single: rtnSingle = targetValue
    
    If targetValue < 0 Then
        rtnSingle = (targetValue * -1) - (referenceValue / 2)
    End If
    
    Position_Centering = rtnSingle
    
End Function


Private Function Calendar_Show(targetLeft As Single, targetTop As Single, targetHeight As Single, targetWidth As Single)

    Dim buffShape As Shape
    
    Dim i As Integer
    Dim cntX As Integer
    Dim cntY As Single
    
    Dim buffHeight As Variant
    Dim buffWidth As Variant
    Dim pitchX As Single
    Dim pitchY As Single
    
    Dim targetDate As Date: targetDate = Get_TargetDate()
    
    Dim arrayWeekDay(6) As String
            arrayWeekDay(0) = "Sun"
            arrayWeekDay(1) = "Mon"
            arrayWeekDay(2) = "Tue"
            arrayWeekDay(3) = "Wed"
            arrayWeekDay(4) = "Thu"
            arrayWeekDay(5) = "Fri"
            arrayWeekDay(6) = "Sat"
             
     
     ' ---- Delete Calendar UI ----
     Call Calendar_Delete
     
     
     ' ---- Calendar Back Ground ----
     Set buffShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, targetLeft, targetTop, targetHeight, targetWidth)
     With buffShape
        .Name = "Calendar_BackGround_(" + ActiveSheet.Name + ")"
        .OnAction = "'Dummy_Function'"
        .Fill.ForeColor.RGB = defaultColor.LightColor
        .Line.Visible = msoFalse
        .Adjustments(1) = 0.03
        With .Shadow
            .Style = msoShadowStyleOuterShadow
            .Transparency = 0.6
            .Size = 102
            .Blur = 5
            .OffsetX = 0
            .OffsetY = 2
            .Visible = msoTrue
            .ForeColor.RGB = shadowColor
        End With
     End With

    
    pitchX = targetHeight / 7
    pitchY = targetWidth / 8
    
    ' ---- Day Button Draw (0,2) ----
    cntX = 0
    cntY = 2
    For i = 1 To 42
        Set buffShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 8, 8, 8, 8)
        With buffShape
            .Name = "Calendar_Button_Day_" + Format(CStr(i), "00")
            .OnAction = "'Calendar_DayClick " + CStr(i) + "'"
            .Adjustments(1) = 0.1
            .TextFrame.Characters.Font.Color = defaultColor.DarkColor
            .Fill.ForeColor.RGB = defaultColor.LightColor
            .Line.ForeColor.RGB = accentColor.DarkColor
            .Line.Visible = msoFalse
            .TextFrame.Characters.Text = Format(CStr(i), "00")
            .TextEffect.FontName = "Meiryo UI"
            .TextFrame.Characters.Font.Size = pitchX / 2 - pitchX / 8           '           : サイズ
            .TextFrame.HorizontalAlignment = xlHAlignCenter                     ' 横位置    : 中央寄せ
            .TextFrame.VerticalAlignment = xlVAlignCenter                       ' 縦位置    : 中央寄せ
            .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow    ' 横オーバーフロー  : True
            .TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow        ' 縦オーバーフロー  : True
            .TextFrame2.WordWrap = msoFalse                                     ' 改行      : False
            .Height = pitchY
            .Width = pitchX
            .Top = targetTop + (pitchY * cntY)
            .Left = targetLeft + (pitchX * cntX)
                    
        End With
        
        If cntX >= 6 Then
            cntX = 0
            cntY = cntY + 1
        Else
            cntX = cntX + 1
        
        End If
    
    Next i
    
    
    ' ---- WeekDay Draw (0,1) ----
    cntX = 0
    For i = 0 To 6
        Set buffShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 8, 8, 8, 8)
        With buffShape
            .Name = "Calendar_Label_Weekday_" + arrayWeekDay(i)
            .OnAction = "'Dummy_Function'"
            .TextFrame.Characters.Font.Color = shadowColor
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = defaultColor.LightColor
            .Line.Visible = msoFalse
            .TextFrame.Characters.Text = arrayWeekDay(i)
            .TextEffect.FontName = "Meiryo UI"
            .TextFrame.Characters.Font.Size = pitchX / 3                        '           : サイズ
            .TextFrame.HorizontalAlignment = xlHAlignCenter                     ' 横位置    : 中央寄せ
            .TextFrame.VerticalAlignment = xlVAlignCenter                       ' 縦位置    : 中央寄せ
            .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow    ' 横オーバーフロー  : True
            .TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow        ' 縦オーバーフロー  : True
            .TextFrame2.WordWrap = msoFalse                                     ' 改行      : False
            .Height = pitchY
            .Width = pitchX
            .Top = targetTop + (pitchY * 1)
            .Left = targetLeft + (pitchX * cntX)
        End With
        
        cntX = cntX + 1
    
    Next i
    
    
    ' ---- Prev. Month Button Draw (1,0) ----
    Set buffShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 8, 8, 8, 8)
    With buffShape
        .Name = "Calendar_Button_PrevMonth"
        .OnAction = "'Calendar_MonthClick" + " " + "-1" + "'"
        .Adjustments(1) = 0.1
        .TextFrame.Characters.Font.Color = accentColor.DarkColor
        .Fill.ForeColor.RGB = accentColor.LightColor
        .Line.ForeColor.RGB = accentColor.DarkColor
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "<"
        .TextFrame.Characters.Font.Bold = True
        .TextEffect.FontName = "Meiryo UI"
        .TextFrame.Characters.Font.Size = pitchX / 2                        '           : サイズ
        .TextFrame.HorizontalAlignment = xlHAlignCenter                     ' 横位置    : 中央寄せ
        .TextFrame.VerticalAlignment = xlVAlignCenter                       ' 縦位置    : 中央寄せ
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow    ' 横オーバーフロー  : True
        .TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow        ' 縦オーバーフロー  : True
        .TextFrame2.WordWrap = msoFalse                                     ' 改行      : False
        .Height = pitchY
        .Width = pitchX
        .Top = targetTop + (pitchY * 0) + (pitchY / 16)
        .Left = targetLeft + (pitchX * 1)
    End With
    
    
    ' ---- Next Month Button Draw (5,0) ----
    Set buffShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 8, 8, 8, 8)
    With buffShape
        .Name = "Calendar_Button_NextMonth"
        .OnAction = "'Calendar_MonthClick" + " " + "1" + "'"
        .Adjustments(1) = 0.1
        .TextFrame.Characters.Font.Color = accentColor.DarkColor
        .Fill.ForeColor.RGB = accentColor.LightColor
        .Line.ForeColor.RGB = accentColor.DarkColor
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = ">"
        .TextFrame.Characters.Font.Bold = True
        .TextEffect.FontName = "Meiryo UI"
        .TextFrame.Characters.Font.Size = pitchX / 2                        '           : サイズ
        .TextFrame.HorizontalAlignment = xlHAlignCenter                     ' 横位置    : 中央寄せ
        .TextFrame.VerticalAlignment = xlVAlignCenter                       ' 縦位置    : 中央寄せ
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow    ' 横オーバーフロー  : True
        .TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow        ' 縦オーバーフロー  : True
        .TextFrame2.WordWrap = msoFalse                                     ' 改行      : False
        .Height = pitchY
        .Width = pitchX
        .Top = targetTop + (pitchY * 0) + (pitchY / 16)
        .Left = targetLeft + (pitchX * 5)
    End With
    
    
    ' ---- Year and Month draw ----
    Set buffShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 8, 8, 8, 8)
    With buffShape
        .Name = "Calendar_TextBox_Date"
        .OnAction = "'Dummy_Function'"
        .Adjustments(1) = 0.3
        .TextFrame.Characters.Font.Color = defaultColor.DarkColor
        .Fill.ForeColor.RGB = defaultColor.LightColor
        .Line.ForeColor.RGB = defaultColor.LightColor
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = CStr(Year(targetDate)) + " . " + CStr(Month(targetDate))
        .TextFrame.Characters.Font.Bold = False
        .TextEffect.FontName = "Meiryo UI"
        .TextFrame.Characters.Font.Size = pitchX / 2 + pitchX / 16           '           : サイズ
        .TextFrame.HorizontalAlignment = xlHAlignCenter                     ' 横位置    : 中央寄せ
        .TextFrame.VerticalAlignment = xlVAlignCenter                       ' 縦位置    : 中央寄せ
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow    ' 横オーバーフロー  : True
        .TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow        ' 縦オーバーフロー  : True
        .TextFrame2.WordWrap = msoFalse                                     ' 改行      : False
        .Height = pitchY
        .Width = pitchX * 3
        .Top = targetTop + (pitchY * 0) + (pitchY / 12)
        .Left = targetLeft + (pitchX * 2)
        
    End With
    
    
    ' ---- Day Drow ----
    Call Calendar_DayDraw(targetDate)
    
    
End Function


Private Function Calendar_DayDraw(targetDate As Date)
    Dim targetShape As Shape
    Dim cntTable As Integer
    
    Dim currentTableDate As Date
        currentTableDate = Get_TableStart_Date(targetDate)
        
    Dim originDate As Date:
        originDate = Get_TargetDate()
        
    For cntTable = 1 To 42
        Set targetShape = ActiveSheet.Shapes("Calendar_Button_Day_" + Format(cntTable, "00"))
        With targetShape
            .TextFrame.Characters.Text = Day(currentTableDate)
            
            If Month(currentTableDate) <> Month(targetDate) Then
                .TextFrame.Characters.Font.Color = shadowColor
            Else
                .TextFrame.Characters.Font.Color = defaultColor.DarkColor
            End If
                        
            If currentTableDate = originDate Then
                .OnAction = "'Calendar_DayClick_2Times " + CStr(cntTable) + "'"
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = accentColor.DarkColor
                .TextFrame.Characters.Font.Color = accentColor.LightColor
                .ZOrder (msoBringToFront)
            Else
                .Line.Visible = msoFalse
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = defaultColor.LightColor
            End If
            
            If currentTableDate = Date Then
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = accentColor.DarkColor
                .ZOrder (msoBringToFront)
            End If
            
        End With
    
    currentTableDate = DateAdd("d", 1, currentTableDate)
    
    Next cntTable
    
End Function


Private Function Get_TableStart_Date(targetDate As Date) As Date
    Dim initialWeekDay As Integer
        initialWeekDay = Weekday(DateSerial(Year(targetDate), Month(targetDate), 1))
    
    Get_TableStart_Date = DateAdd("d", _
                                    (-1 * (initialWeekDay - 1)), _
                                    DateSerial(Year(targetDate), Month(targetDate), 1))
    
End Function




Private Function Get_TargetDate() As Date
    Dim buffStr As String: buffStr = objTargetDate.Text
    
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim targetDay As Integer
        
    targetYear = Year(buffStr)
    targetMonth = Month(buffStr)
    targetDay = Day(buffStr)
    
    Get_TargetDate = DateSerial(targetYear, targetMonth, targetDay)

End Function


Private Function Dummy_Function()
    ' Do not anything.
End Function


Private Function Calendar_Delete()
    Dim bufShape As Shape
    For Each buffShape In ActiveSheet.Shapes
        If buffShape.Name Like ("*Calendar*") Then
            buffShape.Delete
        End If
    Next
    
End Function
