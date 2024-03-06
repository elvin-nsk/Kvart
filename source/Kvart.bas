Attribute VB_Name = "Kvart"
'===============================================================================
'   ������          : Kvart
'   ��������        : ��������� ������� ��������� ����������� �� ������ �������
'   ������          : 2024.03.04
'   �����           : https://vk.com/elvin_macro/Kvart
'                     https://github.com/elvin-nsk/Kvart
'   �����           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Kvart"

'===============================================================================
' # Globals

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
    IsPrevMonth As Boolean
    IsNextMonth As Boolean
    IsSmallMonths As Boolean
    ErrorLog As Logger
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
    PrevMonthRect As Rect
    NextMonthRect As Rect
End Type

'����� �������� � �������
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
Const PREV_MONTH_FRAME As String = "PREV_MONTH_FRAME"
Const NEXT_MONTH_FRAME As String = "NEXT_MONTH_FRAME"

'� ����� ������ ��������� �������� ��������
Const WEEK_FRAME_NUM As String = "1"
Const DAY_FRAME_NUM As String = "2"
Const SUN_FRAME_NUM As String = "7"
Const DAY_DUB_FRAME_NUM As String = "29"
Const SMALLDAY_FRAME_NUM As String = "1"
Const SMALLSUN_FRAME_NUM As String = "6"

'===============================================================================
' # Entry points

Sub Start()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Doc As Document
    With InputData.RequestDocumentOrPage
        If .IsError Then Exit Sub
        Set Doc = .Document
    End With
    
    Dim Params As typeParams: Params = ExtractParamsFromActivePage
    
    If Not ValidateActivePage(Params) Then
        Params.ErrorLog.Check
        Exit Sub
    End If
            
    With FindShapesActivePageLayers(False, True).CreateDocumentFrom
        .Name = "��������� �� ������ ������� " & Doc.Name
        .Activate
    End With
    BoostStart APP_NAME
    MakeKvartFromActiveDoc Params
        
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Main routine

Private Sub MakeKvartFromActiveDoc(ByRef Params As typeParams)
    
    #If DebugMode = 1 Then
    Const CAL_PAGES_COUNT As Long = 3
    #Else
    Const CAL_PAGES_COUNT As Long = 14
    #End If
    
    Dim Positions As typePositions: Positions = _
        CalculatePositionsFromActivePage(Params)
    
    With Params
        .YearKvart = CLng(FindByName(YEAR_NAME).Text.Story.Text)
        
        '��������� ���������� ���� � ������ �������
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
    DuplicateActivePage CAL_PAGES_COUNT - 1
    
    Dim SmallMonthsSources As Collection
    If Params.IsSmallMonths Then Set SmallMonthsSources = New Collection
    Dim PBar As ProgressBar: Set PBar = _
        ProgressBar.New_(ProgressBarNumeric, CAL_PAGES_COUNT)
    PBar.Caption = "���������� �����"
    For i = 1 To CAL_PAGES_COUNT
        ActiveDocument.Pages(i).Activate
        ProcessActivePage Params, Positions
        If Params.IsSmallMonths Then SmallMonthsSources.Add GetMonthSource
        PBar.Update
    Next i
    If Params.IsSmallMonths Then
        PBar.Caption = "����������� �������������� �������"
        For i = 1 To CAL_PAGES_COUNT
            ProcessSmallMonthFrames _
                ActiveDocument.Pages(i), SmallMonthsSources, Params, Positions
            PBar.UpdateTo i
        Next i
    End If
    
    ActiveDocument.Pages(1).Activate

End Sub

'===============================================================================
' # Logic

Private Sub ProcessActivePage( _
                ByRef Params As typeParams, _
                ByRef Positions As typePositions _
            )
    
    Dim Shape As Shape, Shapes As New ShapeRange
    Dim WeekSrc As Shape, DaySrc As Shape
    Dim SunSrc As Shape, DaydubSrc As Shape
    Dim SmalldaySrc As Shape, SmallsunSrc As Shape, FrameSrc As Shape
    Dim Year As Long, Month As Long, Day As Long
    Dim MonthInArr As Long, CurWeek As Long
    Dim StartDay As Long, DubsCount As Long
    Dim i

    Select Case ActivePage.Index
        Case 2 To 13
            Year = Params.YearKvart
            Month = ActivePage.Index - 1
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
    Dim MonthRUtxt As String: MonthRUtxt = VBA.UCase(Params.MonthRU(Month))
    Dim MonthENtxt As String: MonthENtxt = VBA.UCase(Params.MonthEN(Month))
    
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
        DubsCount = _
            Params.MaxFrame - Params.DaysIn(MonthInArr) - (StartDay - 1)
        If DubsCount < 0 Then DubsCount = Abs(DubsCount) Else DubsCount = 0
    Else
        DubsCount = 0
    End If
    
    '����������� ������ ������
    If Params.IsWeeks Then
        Set WeekSrc = FindByName(WEEK_NUM_NAME)
        For i = 1 To Params.MaxWeek
            CurWeek = _
                VBA.DatePart( _
                    "ww", _
                    DateAdd( _
                        "ww", -1 + i, "1/" & CStr(Month) & "/" & CStr(Year) _
                    ), _
                    vbMonday, vbFirstFourDays _
                )
            Set Shape = _
                DupShape( _
                    WeekSrc, FindByName(WEEK_FRAME_PREFIX + CStr(i)), _
                    Positions.WeekNumShiftX, Positions.WeekNumShiftY _
                )
            Shape.Text.Story.Text = CStr(CurWeek)
            Shape.Name = vbNullString
        Next
        TryDeleteByName WEEK_NUM_NAME
    End If
    
    '����������� ���
    For i = 1 To Params.MaxFrame
        Day = i - (StartDay - 1)
        Set FrameSrc = FindByName(DAY_FRAME_PREFIX + CStr(i))
        Select Case True
            Case Params.IsSmalls And i < StartDay And IsSun(Month, Day, i)
                Set Shape = _
                    DupShape( _
                        SmallsunSrc, FrameSrc, _
                        Positions.SmallsunNumShiftX, Positions.SmallsunNumShiftY _
                    )
                Shape.Text.Story.Text = _
                    CStr(Params.DaysIn(MonthInArr - 1) - (StartDay - 1) + i)
                Shape.Name = vbNullString
            Case Params.IsSmalls And i < StartDay
                Set Shape = _
                    DupShape( _
                        SmalldaySrc, FrameSrc, _
                        Positions.SmalldayNumShiftX, Positions.SmalldayNumShiftY _
                    )
                Shape.Text.Story.Text = _
                    CStr(Params.DaysIn(MonthInArr - 1) - (StartDay - 1) + i)
                Shape.Name = vbNullString
            Case Params.IsSmalls _
             And i > Params.DaysIn(MonthInArr) + (StartDay - 1) _
             And IsSun(Month, Day, i)
                Set Shape = _
                    DupShape( _
                        SmallsunSrc, FrameSrc, _
                        Positions.SmallsunNumShiftX, Positions.SmallsunNumShiftY _
                    )
                Shape.Text.Story.Text = _
                    CStr(i - Params.DaysIn(MonthInArr) - (StartDay - 1))
                Shape.Name = vbNullString
            Case Params.IsSmalls _
             And i > Params.DaysIn(MonthInArr) + (StartDay - 1)
                Set Shape = _
                    DupShape( _
                        SmalldaySrc, FrameSrc, _
                        Positions.SmalldayNumShiftX, Positions.SmalldayNumShiftY _
                    )
                Shape.Text.Story.Text = _
                    CStr(i - Params.DaysIn(MonthInArr) - (StartDay - 1))
                Shape.Name = vbNullString
            Case (i = 29 And DubsCount > 0) Or (i = 30 And DubsCount = 2)
                Set Shape = _
                    DupShape( _
                        DaydubSrc, FrameSrc, _
                        Positions.DaydubShiftX, Positions.DaydubShiftY _
                    )
                Shapes.RemoveAll
                Shapes.Add Shape
                Shapes.UngroupAll
                Shapes(NUM_TOP_NAME).Text.Story.Text = _
                    CStr(i - (StartDay - 1))
                Shapes(NUM_BOT_NAME).Text.Story.Text = _
                    CStr(i - (StartDay - 1) + 7)
                Shapes(NUM_TOP_NAME).Name = vbNullString
                Shapes(NUM_BOT_NAME).Name = vbNullString
            Case IsSun(Month, Day, i) _
             And i >= StartDay _
             And i <= Params.DaysIn(MonthInArr) + (StartDay - 1)
                Set Shape = _
                    DupShape( _
                        SunSrc, FrameSrc, _
                        Positions.SunNumShiftX, Positions.SunNumShiftY _
                    )
                Shape.Text.Story.Text = CStr(Day)
                Shape.Name = vbNullString
            Case i >= StartDay _
             And i <= Params.DaysIn(MonthInArr) + (StartDay - 1)
                Set Shape = _
                    DupShape( _
                        DaySrc, FrameSrc, _
                        Positions.DayNumShiftX, Positions.DayNumShiftY _
                    )
                Shape.Text.Story.Text = CStr(Day)
                Shape.Name = vbNullString
        End Select
    Next
    
    '��������� ������ ���������
    TryDeleteByName DAY_NUM_NAME
    TryDeleteByName SUN_NUM_NAME
    TryDeleteByName DAY_DUB_NAME
    TryDeleteByName SMALLDAY_NUM_NAME
    TryDeleteByName SMALLSUN_NUM_NAME
    
    '��������� ������ �����
    If Params.IsDubs Then
        TryDeleteByName WEEK_FRAME_PREFIX & "6"
        For i = 36 To 42
            TryDeleteByName DAY_FRAME_PREFIX & VBA.CStr(i)
        Next
    End If
    
End Sub

'===============================================================================
' # Helpers

Private Property Get GetMonthSource() As ShapeRange
    Dim Shapes As ShapeRange: Set Shapes = ActivePage.SelectableShapes.All
    RemoveSmallMonthFrames Shapes
    Set GetMonthSource = Shapes
End Property

Private Sub ProcessSmallMonthFrames( _
                ByVal TargetPage As Page, _
                ByVal SmallMonthsSources As Collection, _
                ByRef Params As typeParams, _
                ByRef Positions As typePositions _
            )
    If Params.IsPrevMonth Then
        TryCopyShapesToFrame _
            TargetPage.Previous, TargetPage, _
            SmallMonthsSources, Positions.PrevMonthRect
    End If
    If Params.IsNextMonth Then
        TryCopyShapesToFrame _
            TargetPage.Next, TargetPage, _
            SmallMonthsSources, Positions.NextMonthRect
    End If
End Sub

Private Sub TryCopyShapesToFrame( _
                ByVal FromPage As Page, _
                ByVal TargetPage As Page, _
                ByVal SmallMonthsSources As Collection, _
                ByVal FrameRect As Rect _
            )
    If FromPage Is Nothing Then Exit Sub
    Dim Shapes As ShapeRange: Set Shapes = _
        SmallMonthsSources(FromPage.Index).Duplicate
    Shapes.MoveToLayer TargetPage.ActiveLayer
    FitInside Group(Shapes), FrameRect
End Sub

Private Sub RemoveSmallMonthFrames(ByVal FromRange As ShapeRange)
    TryRemoveFromRange FromRange, FromRange.Shapes.FindShape(PREV_MONTH_FRAME)
    TryRemoveFromRange FromRange, FromRange.Shapes.FindShape(NEXT_MONTH_FRAME)
End Sub

Private Sub TryRemoveFromRange(ByVal Range As ShapeRange, ByVal Shape As Shape)
    If Shape Is Nothing Then Exit Sub
    Range.RemoveRange PackShapes(Shape)
End Sub

'���������� �������� ����������
Private Function ExtractParamsFromActivePage() As typeParams
    With ExtractParamsFromActivePage
        .IsWeeks = Not NotFound(WEEK_NUM_NAME)
        .IsDubs = Not NotFound(DAY_DUB_NAME)
        If NotFound(SMALLDAY_NUM_NAME) _
        Or NotFound(SMALLSUN_NUM_NAME) Then _
            .IsSmalls = False Else .IsSmalls = True
        If .IsDubs Then
            .MaxWeek = 5
            .MaxFrame = 35
        Else
            .MaxWeek = 6
            .MaxFrame = 42
        End If
        .IsNextMonth = Not NotFound(NEXT_MONTH_FRAME)
        .IsPrevMonth = Not NotFound(PREV_MONTH_FRAME)
        .IsSmallMonths = .IsNextMonth Or .IsPrevMonth
        
        .MonthRU(1) = "������"
        .MonthRU(2) = "�������"
        .MonthRU(3) = "����"
        .MonthRU(4) = "������"
        .MonthRU(5) = "���"
        .MonthRU(6) = "����"
        .MonthRU(7) = "����"
        .MonthRU(8) = "������"
        .MonthRU(9) = "��������"
        .MonthRU(10) = "�������"
        .MonthRU(11) = "������"
        .MonthRU(12) = "�������"
    
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

'�������� �� ������
Private Function ValidateActivePage(ByRef Params As typeParams) As Boolean
    With Params
        Set .ErrorLog = New Logger
    
        '������ 1-�� ������ (������ �����������)
        LogIfNotFound YEAR_NAME, "�������� ����", Params
        LogIfNotFound MONTH_RU_NAME, "�������� ������ ��-������", Params
        LogIfNotFound DAY_NUM_NAME, "������� ���", Params
        LogIfNotFound SUN_NUM_NAME, "��������� ���", Params
        If .IsDubs Then
            LogIfNotFound NUM_TOP_NAME, "������� ����� �������� ���", Params
            LogIfNotFound NUM_BOT_NAME, "������ ����� �������� ���", Params
        End If
        Dim i As Long
        For i = 1 To .MaxFrame
            LogIfNotFound DAY_FRAME_PREFIX + VBA.CStr(i), "����� ���", Params
        Next
        If .IsWeeks Then
            For i = 1 To .MaxWeek
                LogIfNotFound _
                    WEEK_FRAME_PREFIX + VBA.CStr(i), _
                    "����� ������ ������", Params
            Next
        End If
        If .ErrorLog.Count > 0 Then Exit Function
        
        '������ 2-�� ������ (������ �� ���������)
        LogIfNotText YEAR_NAME, "�������� ����", Params
        LogIfNotText MONTH_RU_NAME, "�������� ������ ��-������", Params
        LogIfNotText DAY_NUM_NAME, "������� ���", Params
        LogIfNotText SUN_NUM_NAME, "��������� ���", Params
        If .IsDubs Then
            LogIfNotText NUM_TOP_NAME, "������� ����� �������� ���", Params
            LogIfNotText NUM_BOT_NAME, "������ ����� �������� ���", Params
        End If
        If .ErrorLog.Count > 0 Then Exit Function
        
        '������ 3-�� ������ (����� � ������� �� �����)
        LogIfNotNum YEAR_NAME, "�������� ����", Params
        If .ErrorLog.Count > 0 Then Exit Function
    End With
    
    ValidateActivePage = True
End Function

'���������� ���������� � ������������
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
        If Params.IsPrevMonth Then _
            Set .PrevMonthRect = _
                FindByName(PREV_MONTH_FRAME).BoundingBox.GetCopy
        If Params.IsNextMonth Then _
            Set .NextMonthRect = _
                FindByName(NEXT_MONTH_FRAME).BoundingBox.GetCopy
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

Private Sub TryDeleteByName(ByVal Name As String)
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

Private Sub LogIfNotFound( _
                Name As String, _
                objText As String, _
                Params As typeParams _
            )
    If NotFound(Name) Then _
        Params.ErrorLog.Add "�� ������ ������ " & objText & " (" & Name & ")"
End Sub

Private Sub LogIfNotNum( _
                Name As String, _
                objText As String, _
                Params As typeParams _
            )
    Dim Shape As Shape
    If NotNum(Name, Shape) Then _
        Params.ErrorLog.Add _
            "����� � ������� " & objText & " (" & YEAR_NAME & ")" _
          & " �� �������� ������", Shape
End Sub

Private Sub LogIfNotText( _
                Name As String, _
                objText As String, _
                Params As typeParams _
            )
    Dim Shape As Shape
    If NotText(Name, Shape) Then _
        Params.ErrorLog.Add _
            "������ " & objText & " (" & Name & ")" & " - �� ���������", _
            Shape
End Sub

Private Property Get NotFound(ByVal Name As String) As Boolean
    NotFound = FindByName(Name) Is Nothing
End Property

Private Property Get NotNum( _
                         ByVal Name As String, _
                         Optional ByRef ReturnShape As Shape _
                     ) As Boolean
    Set ReturnShape = FindByName(Name)
    If ReturnShape Is Nothing Then
        NotNum = True
        Exit Property
    End If
    NotNum = Not VBA.IsNumeric(ReturnShape.Text.Story.Text)
End Property

Private Property Get NotText( _
                         ByVal Name As String, _
                         Optional ByRef ReturnShape As Shape _
                     ) As Boolean
    Set ReturnShape = FindByName(Name)
    If ReturnShape Is Nothing Then
        NotText = True
        Exit Property
    End If
    NotText = Not (ReturnShape.Type = cdrTextShape)
End Property

'===============================================================================
' # �����

Private Sub testSomething()
'
End Sub
