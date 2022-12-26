Attribute VB_Name = "Kvart"
'===============================================================================
'   Макрос          : Kvart
'   Описание        : Генератор месяцев календаря постранично на основе шаблона
'   Версия          : 2022.12.26
'   Сайты           : https://vk.com/elvin_macro/Kvart
'                     https://github.com/elvin-nsk/Kvart
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

'===============================================================================

Public Type typeParams
    MonthRU(1 To 12) As String
    MonthEN(1 To 12) As String
    DaysIn(-1 To 13) As Long
    YearKvart As Long
    MaxWeek As Long
    MaxFrame As Long
    IsDubs As Boolean
    IsWeeks As Boolean
    IsSmalls As Boolean
    ErrLog As Logger
End Type

Public Type typePositions
    WeekNumShiftX As Double
    WeekNumShiftY As Double
    DayNumShiftX As Double
    DayNumShiftY As Double
    SunNumShiftX As Double
    SunNumShiftY As Double
    DaydubShiftX As Double
    DaydubShiftY As Double
    SmalldayNumShiftX As Double
    SmalldayNumShiftY As Double
    SmallsunNumShiftX As Double
    SmallsunNumShiftY As Double
End Type

'имена объектов в шаблоне
Const YEAR_NAME As String = "YEAR"
Const MONTH_RU_NAME As String = "MONTH_RU"
Const MONTH_EN_NAME As String = "MONTH_EN"
Const MONTH_NUM As String = "MONTH_NUM"
Const WEEK_NUM_NAME As String = "WEEK_NUM"
Const DAY_NUM_NAME As String = "DAY_NUM"
Const SUN_NUM_NAME As String = "SUN_NUM"
Const DAY_DUB_NAME As String = "DAY_DUB"
Const NUM_TOP_NAME As String = "NUM_TOP"
Const NUM_BOT_NAME As String = "NUM_BOT"
Const SMALLDAY_NUM_NAME As String = "SMALLDAY_NUM"
Const SMALLSUN_NUM_NAME As String = "SMALLSUN_NUM"
Const DAY_FRAME_PREFIX As String = "DAY_FRAME_"
Const WEEK_FRAME_PREFIX As String = "WEEK_FRAME_"

'к каким рамкам привязаны ключевые элементы
Const WEEK_FRAME_NUM As String = "1"
Const DAY_FRAME_NUM As String = "2"
Const SUN_FRAME_NUM As String = "7"
Const DAY_DUB_FRAME_NUM As String = "29"
Const SMALLDAY_FRAME_NUM As String = "1"
Const SMALLSUN_FRAME_NUM As String = "6"

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    Dim Doc As Document
    With InputData.GetDocumentOrPage
        If .IsError Then Exit Sub
        Set Doc = .Document
    End With
    
    Dim Params As typeParams
    Params = ExtractParamsFromActivePage
    
    If Not ValidateActivePage(Params) Then
        Params.ErrLog.Check
        Exit Sub
    End If
            
    With FindShapesActivePageLayers(False, True).CreateDocumentFrom
        .Name = "Календарь на основе шаблона " & Doc.Name
        .Activate
    End With
    BoostStart "Kvart", RELEASE
    MakeKvartFromActiveDoc Params
        
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================

Private Sub MakeKvartFromActiveDoc(ByRef Params As typeParams)
    
    Const CalPagesCount As Long = 14
    
    Dim Positions As typePositions
    Positions = CalculatePositionsFromActivePage(Params)
    
    With Params
        .YearKvart = VBA.CLng(FindByName(YEAR_NAME).Text.Story.Text)
        
        'вычисляем количество дней в нужных месяцах
        Dim i As Long
        For i = 1 To 12
            .DaysIn(i) = DaysInMonth(.YearKvart, i)
        Next
        .DaysIn(-1) = DaysInMonth(.YearKvart - 1, 11)
        .DaysIn(0) = DaysInMonth(.YearKvart - 1, 12)
        .DaysIn(13) = DaysInMonth(.YearKvart + 1, 1)
    End With
    
    ActiveDocument.MasterPage.SetSize _
        ActivePage.SizeWidth, _
        ActivePage.SizeHeight
    DuplicateActivePage CalPagesCount - 1
    
    Dim PBar As IProgressBar
    Set PBar = ProgressBar.CreateNumeric(CalPagesCount)
    PBar.Caption = "Заполнение сеток"
    For i = 1 To VBA.IIf(RELEASE, CalPagesCount, 2)
        ProcessPage ActiveDocument.Pages(i), Params, Positions
        PBar.Update
    Next
    
    ActiveDocument.Pages(1).Activate

End Sub

Private Sub ProcessPage( _
                ByVal Page As Page, _
                ByRef Params As typeParams, _
                ByRef Positions As typePositions _
            )
    
    Page.Activate
    
    Dim Shape As Shape, Shapes As New ShapeRange
    Dim WeekSrc As Shape, DaySrc As Shape
    Dim SunSrc As Shape, DaydubSrc As Shape
    Dim SmalldaySrc As Shape, SmallsunSrc As Shape, FrameSrc As Shape
    Dim MonthRUtxt As String, MonthENtxt As String
    Dim Year As Long, Month As Long, Day As Long
    Dim MonthInArr As Long, CurWeek As Long
    Dim StartDay As Long, DubsCount As Long
    Dim i

    Select Case Page.Index
        Case 2 To 13
            Year = Params.YearKvart
            Month = Page.Index - 1
            MonthInArr = Month
        Case 1
            Year = Params.YearKvart - 1
            Month = 12
            MonthInArr = 0
        Case 14
            Year = Params.YearKvart + 1
            Month = 1
            MonthInArr = 13
    End Select
    MonthRUtxt = VBA.UCase(Params.MonthRU(Month))
    MonthENtxt = VBA.UCase(Params.MonthEN(Month))
    
    SetTextByName MONTH_RU_NAME, MonthRUtxt
    SetTextByName MONTH_EN_NAME, MonthENtxt
    SetTextByName MONTH_NUM, Month
    SetTextByName YEAR_NAME, Year
    
    StartDay = _
        VBA.DatePart( _
            "w", "1/" & CStr(Month) & "/" + CStr(Year), _
            vbMonday, vbFirstFourDays _
        )
    Set DaySrc = FindByName(DAY_NUM_NAME)
    Set SunSrc = FindByName(SUN_NUM_NAME)
    Set DaydubSrc = FindByName(DAY_DUB_NAME)
    
    If Params.IsSmalls Then
        Set SmalldaySrc = FindByName(SMALLDAY_NUM_NAME)
        Set SmallsunSrc = FindByName(SMALLSUN_NUM_NAME)
    End If
    If Params.IsDubs Then
        DubsCount = Params.MaxFrame - Params.DaysIn(MonthInArr) - (StartDay - 1)
        If DubsCount < 0 Then DubsCount = Abs(DubsCount) Else DubsCount = 0
    Else
        DubsCount = 0
    End If
    
    'расставляем номера недель
    If Params.IsWeeks Then
        Set WeekSrc = FindByName(WEEK_NUM_NAME)
        For i = 1 To Params.MaxWeek
            CurWeek = VBA.DatePart("ww", DateAdd("ww", -1 + i, "1/" & CStr(Month) & "/" & CStr(Year)), vbMonday, vbFirstFourDays)
            Set Shape = DupShape(WeekSrc, FindByName(WEEK_FRAME_PREFIX + CStr(i)), Positions.WeekNumShiftX, Positions.WeekNumShiftY)
            Shape.Text.Story.Text = CStr(CurWeek)
            Shape.Name = ""
        Next
        SafeDeleteByName WEEK_NUM_NAME
    End If
    
    'расставляем дни
    For i = 1 To Params.MaxFrame
        Day = i - (StartDay - 1)
        Set FrameSrc = FindByName(DAY_FRAME_PREFIX + CStr(i))
        Select Case True
            Case Params.IsSmalls And i < StartDay And IsSun(Month, Day, i)
                Set Shape = DupShape(SmallsunSrc, FrameSrc, Positions.SmallsunNumShiftX, Positions.SmallsunNumShiftY)
                Shape.Text.Story.Text = CStr(Params.DaysIn(MonthInArr - 1) - (StartDay - 1) + i)
                Shape.Name = ""
            Case Params.IsSmalls And i < StartDay
                Set Shape = DupShape(SmalldaySrc, FrameSrc, Positions.SmalldayNumShiftX, Positions.SmalldayNumShiftY)
                Shape.Text.Story.Text = CStr(Params.DaysIn(MonthInArr - 1) - (StartDay - 1) + i)
                Shape.Name = ""
            Case Params.IsSmalls And i > Params.DaysIn(MonthInArr) + (StartDay - 1) And IsSun(Month, Day, i)
                Set Shape = DupShape(SmallsunSrc, FrameSrc, Positions.SmallsunNumShiftX, Positions.SmallsunNumShiftY)
                Shape.Text.Story.Text = CStr(i - Params.DaysIn(MonthInArr) - (StartDay - 1))
                Shape.Name = ""
            Case Params.IsSmalls And i > Params.DaysIn(MonthInArr) + (StartDay - 1)
                Set Shape = DupShape(SmalldaySrc, FrameSrc, Positions.SmalldayNumShiftX, Positions.SmalldayNumShiftY)
                Shape.Text.Story.Text = CStr(i - Params.DaysIn(MonthInArr) - (StartDay - 1))
                Shape.Name = ""
            Case (i = 29 And DubsCount > 0) Or (i = 30 And DubsCount = 2)
                Set Shape = DupShape(DaydubSrc, FrameSrc, Positions.DaydubShiftX, Positions.DaydubShiftY)
                Shapes.RemoveAll
                Shapes.Add Shape
                Shapes.UngroupAll
                Shapes(NUM_TOP_NAME).Text.Story.Text = CStr(i - (StartDay - 1))
                Shapes(NUM_BOT_NAME).Text.Story.Text = CStr(i - (StartDay - 1) + 7)
                Shapes(NUM_TOP_NAME).Name = ""
                Shapes(NUM_BOT_NAME).Name = ""
            Case IsSun(Month, Day, i) And i >= StartDay And i <= Params.DaysIn(MonthInArr) + (StartDay - 1)
                Set Shape = DupShape(SunSrc, FrameSrc, Positions.SunNumShiftX, Positions.SunNumShiftY)
                Shape.Text.Story.Text = CStr(Day)
                Shape.Name = ""
            Case i >= StartDay And i <= Params.DaysIn(MonthInArr) + (StartDay - 1)
                Set Shape = DupShape(DaySrc, FrameSrc, Positions.DayNumShiftX, Positions.DayNumShiftY)
                Shape.Text.Story.Text = CStr(Day)
                Shape.Name = ""
        End Select
    Next
    
    'подчищаем лишние исходники
    SafeDeleteByName DAY_NUM_NAME
    SafeDeleteByName SUN_NUM_NAME
    SafeDeleteByName DAY_DUB_NAME
    SafeDeleteByName SMALLDAY_NUM_NAME
    SafeDeleteByName SMALLSUN_NUM_NAME
    
    'подчищаем лишние рамки
    If Params.IsDubs Then
        SafeDeleteByName WEEK_FRAME_PREFIX & "6"
        For i = 36 To 42
            SafeDeleteByName DAY_FRAME_PREFIX & VBA.CStr(i)
        Next
    End If
    
End Sub

'извлечение основных параметров
Private Function ExtractParamsFromActivePage() As typeParams
    With ExtractParamsFromActivePage
        If FindByName(WEEK_NUM_NAME) Is Nothing Then _
            .IsWeeks = False Else .IsWeeks = True
        If FindByName(DAY_DUB_NAME) Is Nothing Then _
            .IsDubs = False Else .IsDubs = True
        If FindByName(SMALLDAY_NUM_NAME) Is Nothing _
        Or FindByName(SMALLSUN_NUM_NAME) Is Nothing Then _
            .IsSmalls = False Else .IsSmalls = True
        If .IsDubs Then
            .MaxWeek = 5
            .MaxFrame = 35
        Else
            .MaxWeek = 6
            .MaxFrame = 42
        End If
        
        .MonthRU(1) = "январь"
        .MonthRU(2) = "февраль"
        .MonthRU(3) = "март"
        .MonthRU(4) = "апрель"
        .MonthRU(5) = "май"
        .MonthRU(6) = "июнь"
        .MonthRU(7) = "июль"
        .MonthRU(8) = "август"
        .MonthRU(9) = "сентябрь"
        .MonthRU(10) = "октябрь"
        .MonthRU(11) = "ноябрь"
        .MonthRU(12) = "декабрь"
    
        .MonthEN(1) = "january"
        .MonthEN(2) = "february"
        .MonthEN(3) = "march"
        .MonthEN(4) = "april"
        .MonthEN(5) = "may"
        .MonthEN(6) = "june"
        .MonthEN(7) = "july"
        .MonthEN(8) = "august"
        .MonthEN(9) = "september"
        .MonthEN(10) = "october"
        .MonthEN(11) = "november"
        .MonthEN(12) = "december"
    End With
End Function

'проверка на ошибки
Private Function ValidateActivePage(ByRef Params As typeParams) As Boolean
    With Params
        Set .ErrLog = New Logger
    
        'ошибки 1-го уровня (объект отсутствует)
        CheckNotFound YEAR_NAME, "текущего года", Params
        CheckNotFound MONTH_RU_NAME, "названия месяца по-русски", Params
        CheckNotFound DAY_NUM_NAME, "буднего дня", Params
        CheckNotFound SUN_NUM_NAME, "выходного дня", Params
        If .IsDubs Then
            CheckNotFound NUM_TOP_NAME, "верхней части дробного дня", Params
            CheckNotFound NUM_BOT_NAME, "нижней части дробного дня", Params
        End If
        Dim i As Long
        For i = 1 To .MaxFrame
            CheckNotFound DAY_FRAME_PREFIX + VBA.CStr(i), "рамки дня", Params
        Next
        If .IsWeeks Then
            For i = 1 To .MaxWeek
                CheckNotFound WEEK_FRAME_PREFIX + VBA.CStr(i), "рамки номера недели", Params
            Next
        End If
        If .ErrLog.Count > 0 Then Exit Function
        
        'ошибки 2-го уровня (объект не текстовый)
        CheckNotText YEAR_NAME, "текущего года", Params
        CheckNotText MONTH_RU_NAME, "названия месяца по-русски", Params
        CheckNotText DAY_NUM_NAME, "буднего дня", Params
        CheckNotText SUN_NUM_NAME, "выходного дня", Params
        If .IsDubs Then
            CheckNotText NUM_TOP_NAME, "верхней части дробного дня", Params
            CheckNotText NUM_BOT_NAME, "нижней части дробного дня", Params
        End If
        If .ErrLog.Count > 0 Then Exit Function
        
        'ошибки 3-го уровня (текст в объекте не число)
        CheckNotNum YEAR_NAME, "текущего года", Params
        If .ErrLog.Count > 0 Then Exit Function
    End With
    
    ValidateActivePage = True
End Function

'извлечение переменных и расположений
Private Function CalculatePositionsFromActivePage( _
                     ByRef Params As typeParams _
                 ) As typePositions
    With CalculatePositionsFromActivePage
        .DayNumShiftX = _
            FindByName(DAY_NUM_NAME).LeftX _
          - FindByName(DAY_FRAME_PREFIX & DAY_FRAME_NUM).LeftX
        .DayNumShiftY = _
            FindByName(DAY_NUM_NAME).BottomY _
          - FindByName(DAY_FRAME_PREFIX & DAY_FRAME_NUM).BottomY
        .SunNumShiftX = _
            FindByName(SUN_NUM_NAME).LeftX _
          - FindByName(DAY_FRAME_PREFIX & SUN_FRAME_NUM).LeftX
        .SunNumShiftY = _
            FindByName(SUN_NUM_NAME).BottomY _
          - FindByName(DAY_FRAME_PREFIX & SUN_FRAME_NUM).BottomY
        If Params.IsWeeks Then
            .WeekNumShiftX = _
                FindByName(WEEK_NUM_NAME).LeftX _
              - FindByName(WEEK_FRAME_PREFIX & WEEK_FRAME_NUM).LeftX
            .WeekNumShiftY = _
                FindByName(WEEK_NUM_NAME).BottomY _
              - FindByName(WEEK_FRAME_PREFIX & WEEK_FRAME_NUM).BottomY
        End If
        If Params.IsDubs Then
            .DaydubShiftX = _
                FindByName(DAY_DUB_NAME).LeftX _
              - FindByName(DAY_FRAME_PREFIX & DAY_DUB_FRAME_NUM).LeftX
            .DaydubShiftY = _
                FindByName(DAY_DUB_NAME).BottomY _
              - FindByName(DAY_FRAME_PREFIX & DAY_DUB_FRAME_NUM).BottomY
        End If
        If Params.IsSmalls Then
            .SmalldayNumShiftX = _
                FindByName(SMALLDAY_NUM_NAME).LeftX _
              - FindByName(DAY_FRAME_PREFIX & SMALLDAY_FRAME_NUM).LeftX
            .SmalldayNumShiftY = _
                FindByName(SMALLDAY_NUM_NAME).BottomY _
              - FindByName(DAY_FRAME_PREFIX & SMALLDAY_FRAME_NUM).BottomY
            .SmallsunNumShiftX = _
                FindByName(SMALLSUN_NUM_NAME).LeftX _
              - FindByName(DAY_FRAME_PREFIX & SMALLSUN_FRAME_NUM).LeftX
            .SmallsunNumShiftY = _
                FindByName(SMALLSUN_NUM_NAME).BottomY _
              - FindByName(DAY_FRAME_PREFIX & SMALLSUN_FRAME_NUM).BottomY
        End If
    End With
End Function

Private Function DupShape( _
                     ByRef Src As Shape, _
                     ByRef Frame As Shape, _
                     ByVal ShiftX As Double, _
                     ByVal ShiftY As Double _
                 ) As Shape
    Set DupShape = Src.Duplicate
    DupShape.LeftX = Frame.LeftX + ShiftX
    DupShape.BottomY = Frame.BottomY + ShiftY
End Function

Private Function FindByName(ByVal Name As String) As Shape
    Dim Shapes As ShapeRange
    
    Set Shapes = ActivePage.FindShapes(Name)
    If Shapes.Count > 0 Then
        Set FindByName = Shapes(1)
    End If
End Function

Private Sub SetTextByName(ByVal Name As String, ByVal Text As String)
    Dim Shape As Shape
    
    If FindByName(Name) Is Nothing Then Exit Sub
    
    Set Shape = FindByName(Name)
    If Shape.Type = cdrTextShape Then
        If IsUpperCase(Shape.Text.Story.Text) Then
            Shape.Text.Story.Text = VBA.UCase(Text)
        ElseIf IsLowerCase(Shape.Text.Story.Text) Then
            Shape.Text.Story.Text = VBA.LCase(Text)
        Else
            Shape.Text.Story.Text = Text
            Shape.Text.Story.ChangeCase cdrTextSentenceCase
        End If
    End If
    
End Sub

Private Sub SafeDeleteByName(ByVal Name As String)
    Dim Shape As Shape, Shapes As ShapeRange
    
    Set Shapes = ActivePage.FindShapes(Name)
    If Shapes.Count > 0 Then
        For Each Shape In Shapes
            If Shape Is Nothing Then
            Else
                Shape.Delete
            End If
        Next
    End If
End Sub

Private Function DaysInMonth(y, m) As Long
    Dim d As String
    d = "1/" & CStr(m) & "/" & CStr(y)
    DaysInMonth = DateDiff("d", d, DateAdd("m", 1, d))
End Function

Private Function IsSun(Month, Day, Frame) As Boolean
    Select Case Frame
        Case 6, 7, 13, 14, 20, 21, 27, 28, 34, 35, 41, 42: IsSun = True
    End Select
    If (Month = 1 And Day <= 8 And Day > 0) _
    Or (Month = 2 And Day = 23) _
    Or (Month = 3 And Day = 8) _
    Or (Month = 5 And Day = 1) _
    Or (Month = 5 And Day = 9) _
    Or (Month = 6 And Day = 12) _
    Or (Month = 11 And Day = 4) Then
        IsSun = True
    End If
End Function

Private Sub CheckNotFound( _
                Name As String, _
                objText As String, _
                Params As typeParams _
            )
    If FindByName(Name) Is Nothing Then _
        Params.ErrLog.Add "Не найден объект " & objText & " (" & Name & ")"
End Sub

Private Sub CheckNotText( _
                Name As String, _
                objText As String, _
                Params As typeParams _
            )
    Dim Shape As Shape
    Set Shape = FindByName(Name)
    If Not Shape.Type = cdrTextShape Then _
        Params.ErrLog.Add _
            "Объект " & objText & " (" & Name & ")" & " - не текстовый", _
            Shape
End Sub

Private Sub CheckNotNum( _
                Name As String, _
                objText As String, _
                Params As typeParams _
            )
    Dim Shape As Shape
    Set Shape = FindByName(Name)
    Dim Str As String
    Str = Shape.Text.Story.Text
    If Not VBA.IsNumeric(Str) Then _
        Params.ErrLog.Add _
            "Текст в объекте " & objText & " (" & YEAR_NAME & ")" _
          & " не является числом", Shape
End Sub

'===============================================================================
' # тесты

Private Sub testSomething()
'
End Sub
