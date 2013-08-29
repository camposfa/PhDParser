Attribute VB_Name = "Parser"
Option Explicit
' ---------------- Declarations for Public variables ----------------

' These are arrays for storing codes of various types
' The arrays are initialized in function "Read_Codes"
' All are Public at the module level, i.e. available from any procedure within the module
Public ObserverCodes() As String
Public GroupCodes() As String
Public TreeCodes() As String
Public FoodCodes() As String
Public GroupScanCodes() As String
Public HeaderCodes() As String
Public RangingCodes() As String
Public ClimateCodes() As String
Public VertebrateCodes() As String
Public PointSampleCodes() As String
Public LevelCodes() As String
Public ActivityCodes() As String
Public FoodPatchCodes() As String
Public SelfDirectedCodes() As String
Public EatCodes() As String
Public PostureCodes() As String
Public SubstrateCodes() As String
Public AlarmCodes() As String
Public FollowStatusCodes() As String
Public FollowPartCodes() As String
Public MonkeyCodes() As String
Public GPSUnits() As String
Public CentralityCodes() As String

' This array holds the names of the worksheets (tables) produced by the parser
' It is Public at the module level, i.e. available from any procedure within the module
Public WorksheetNames() As String


Public CurrentGPSColor As String

' -------------- Definitions for custom data types ------------------------

' The InputLine data type stores basic info contained in each line of unparsed data
Public Type InputLine
    LineNum As Long                 ' Count
    Datim As Date
    Data As String                  ' Raw Psion input
    LineType As Integer             ' Possible values are listed in the LineTypeEnum enumeration
    ErrorMessage As String          ' Describes input errors
End Type

' Output type for the Observation table
Public Type ObservationOutputType
    ID As Long
    FocalGroup As String
    Observer As String
    StartObservation As Date
    EndObservation As Date
    DurationOfObservation As Long
    FindType As String
    FindPointID As String
    LeaveType As String
    LeavePointID As String
    IsFullDay As Boolean
End Type

' Output type for GroupScan table
Public Type GroupScanOutputType
    ID As Long
    ObservationID As Long
    WaypointID As String
    ScanSeqNum As Integer
    Datim As Date
    GroupActivity As String
    SpeciesCode As String
    ForestLevel As String
    CanopyHeight As Integer
    GroupHeight As Integer
    Climate As String
    Stage As Integer
End Type

' Output type for Vertebrate table
Public Type VertebrateOutputType
    ID As Long
    GroupScanID As Long
    VertSeqNum As Integer
    Species As String
End Type

' Output type for Follow table
Public Type FollowOutputType
    ID As Long
    ObservationID As Long
    SeqNum As Integer
    FocalAnimal As String
    FollowType As String
    WaypointID As String
    StartFollow As Date
    EndFollow As Date
    DurationOfFollow As Long
    SpeciesCode As String
    IsFollowGood As Boolean
    IsTrackGood As Boolean
    IsPointGood As Boolean
    IsActivityGood As Boolean
    IsForagingGood As Boolean
    IsNoMovement As Boolean
    GPSColor As String
    Error1 As Integer
    Error2 As Integer
    EatTotal As Integer
    AbortType As String
    Comment As String
End Type

' Output type for FollowBlock table
Public Type FollowBlockOutputType
    ID As Long
    FollowID As Long
    StartBlock As Date
    EndBlock As Date
    DurationOfBlock As Long
    IsInTrack As Boolean
    IsInForaging As Boolean
    IsInActivity As Boolean
    IsInPoint As Boolean
End Type

' Output type for the PointSample table
Public Type PointSampleOutputType
    ID As Long
    FollowID As Long
    FollowBlockID As Long
    SeqNum As Integer
    Datim As Date
    StateBehav As String
    SpeciesCode As String
    Posture As String
    Substrate As String
    ForestLevel As String
    Height As Integer
    Centrality As String
    IsCarryingDorsal As Boolean
    NumNeighbors0 As Integer
    NumNeighbors2 As Integer
    NumNeighbors5 As Integer
End Type

' Output type for the Activity table
Public Type ActivityOutputType
    ID As Long
    FollowID As Long
    FollowBlockID As Long
    Activity As String
    StartState As Date
    EndState As Date
    DurationOfState As Long
End Type

' Output type for the SelfDirected table
Public Type SelfDirectedOutputType
    ID As Long
    FollowID As Long
    FollowBlockID As Long
    ActivityID As Long
    Datim As Date
    SeqNum As Integer
    Behavior As String
End Type

' Output type for the FoodPatch table
Public Type FoodPatchOutputType
    ID As Long
    FollowID As Long
    FollowBlockID As Long
    EnterTime As Date
    ExitTime As Date
    PatchType As String
    SpeciesCode As String
    PatchDuration As Long
    IsComplete As Boolean
End Type

' Output type for the FoodObject table
Public Type FoodObjectOutputType
    ID As Long
    FollowID As Long
    FollowBlockID As Long
    FoodPatchID As Long
    FoodItem As String
    SpeciesCode As String
    StartFeeding As Date
    EndFeeding As Date
    DurationOfFeeding As Long
End Type

' Output type for the ForagingEvent table
Public Type ForagingEventOutputType
    ID As Long
    FoodObjectID As Long
    Datim As Date
    SeqNum As Integer
    FoodAction As String
End Type

' Output type for the FruitVisit table
Public Type FruitVisitOutputType
    ID As Long
    ObservationID As Long
    TreeID As String
    WaypointID As String
    SeqNum As Long
    Datim As Date
    SpeciesCode As String
    NumMonkeys As Integer
    LeafCover As Integer
    LeafMaturity As Integer
    FruitCover As Integer
    FruitMaturity As Integer
    FlowerCover As Integer
    FlowerMaturity As Integer
    NumPlants As Integer
    NumFruiting As Integer
End Type

' Output type for the TreeCBH table
Public Type TreeCBHOutputType
    ID As Long
    TreeID As Long
    FruitVisitID As Long
    SeqNum As Integer
    StemNum As Integer
    CBH As Integer
End Type

' Output type for the Alarm table
Public Type AlarmOutputType
    ID As Long
    ObservationID As Long
    WaypointID As String
    Datim As Date
    PredatorType As String
    PredatorSpecies As String
    NumAlarmers As String
    NumAlarms As String
    AlarmerAge As String
    IsConfirmed As Boolean
    Danger As String
    IsMultiple As Boolean
    IsPresent As Boolean
    ForestLevel As String
    Height As Integer
End Type

' Output type for the Interaction table
Public Type InteractionOutputType
    ID As Long
    ObservationID As Long
    Datim As Date
    Actor As String
    Recipient As String
    InteractionType As String
End Type

' Output type for the Intergroup table
Public Type IntergroupOutputType
    ID As Long
    ObservationID As Long
    WaypointID As String
    Datim As Date
    OpponentGroup As String
    Outcome As String
End Type

' Output type for the RangingEvent table (water and vertebrates)
Public Type RangingEventOutputType
    ID As Long
    ObservationID As Long
    WaypointID As String
    Datim As Date
    EventType As String
End Type

Public Type CommentOutputType
    ID As Long
    ObservationID As Long
    Datim As Date
    Comment As String
End Type

' -------------- Enumerations ------------------------

Public Enum LineTypeEnum
    
    ' Tree codes
    TreeID = 1
    TreeWaypoint = 2
    TreeSpecies = 3
    TreeNum = 4
    TreeCBH = 5
    TreePhenology = 6
    TreeBromeliads = 7
    TreeDisks = 8
    TreeEnd = 9
    
    'Group scan codes
    GSWaypoint = 10
    GSClimate = 11
    GSLevel = 12
    GSActivity = 13
    GSHeight = 14
    GSVertebrate = 15
    GSStage = 16
    GSEnd = 17
    
    ' Header codes
    HeaderObserver = 18
    HeaderGroup = 19
    
    ' Ranging codes
    RangingWake = 20
    RangingSleep = 21
    RangingWater = 22
    RangingFind = 23
    RangingLeave = 24
    RangingVertebrate = 25
    
    ' Point sample codes
    PSActivity = 26
    PSPosture = 27
    PSSubstrate = 28
    PSLevel = 29
    PSCentrality = 30
    PSNeighbors = 31
    PSWaypoint = 32
        
    ' Follow activity
    FollowActivity = 33
    
    ' Follow events
    SelfDirected = 34
    EatNew = 35
    EatSame = 36
    EatTotal = 37
    
    ' Foraging codes
    FoodPatchEnter = 38
    FoodPatchEnd = 39
        
    ' Alarm codes
    Alarm = 40
    AlarmWaypoint = 41
    AlarmIntensity = 42
    AlarmLevel = 43
    AlarmDanger = 44
    AlarmMultiple = 45
    AlarmPresent = 46
    AlarmSpecies = 47
    AlarmEnd = 48
    
    ' Behavior codes
    Behavior = 49
    BehaviorIntergroup = 50
    
    ' Follow codes
    FollowStatus = 51
    FollowBlockStatus = 52
    FollowNoMovement = 53
    FollowStart = 54
    FollowEnd = 55
    FollowWaypoint = 56
    Abort = 57
    GPSError = 58
    GPSColor = 59
    
    ' Other
    Comment = 60
    Done = 61
    Other = 62
    Unknown = 63
    Blank = 64
    
End Enum

Public Enum FoodObjectEnum
    Ant = 1
    Bromeliad = 2
    Caterpillar = 3
    Egg = 4
    Fruit = 5
    Insect = 6
    Leaf = 7
    Nest = 8
    FoodOther = 9
    Pith = 10
    Flower = 11
    Seed = 12
    Thorn = 13
    Vertebrate = 14
    Water = 15
End Enum

Public Enum VisibilityEnum
    V_Forage = 1
    V_Activity = 2
    V_PointSample = 3
    V_Track = 4
End Enum


' This is the first Sub called by the program and it controls the sequence of steps for the entire program
Public Sub ProcessData()

    '  Array that stores every line of input
    Dim InputData() As InputLine
   
   ' One variable for each of the output data types
    Dim CurrentObservation As ObservationOutputType
    Dim CurrentGroupScan As GroupScanOutputType
    Dim CurrentVertebrate As VertebrateOutputType
    Dim CurrentFollow As FollowOutputType
    Dim CurrentFollowBlock As FollowBlockOutputType
    Dim CurrentPointSample As PointSampleOutputType
    Dim CurrentActivity As ActivityOutputType
    Dim CurrentSelfDirected As SelfDirectedOutputType
    Dim CurrentFoodPatch As FoodPatchOutputType
    Dim CurrentFoodObject As FoodObjectOutputType
    Dim CurrentForagingEvent As ForagingEventOutputType
    Dim CurrentFruitVisit As FruitVisitOutputType
    Dim CurrentTreeCBH As TreeCBHOutputType
    Dim CurrentAlarm As AlarmOutputType
    Dim CurrentInteraction As InteractionOutputType
    Dim CurrentIntergroup As IntergroupOutputType
    Dim CurrentRangingEvent As RangingEventOutputType
    Dim CurrentComment As CommentOutputType
    
    Dim VertebrateData() As VertebrateOutputType
    Dim TreeCBHData() As Integer
    
    Dim NextNewVertebrateID As Long
    
    ' Counters
    Dim i As Integer
    Dim RowIn As Long
    Dim lastRow As Long
    
    ' These true/false variables keep track of your follow, points sample status
    Dim IsInFollow As Boolean
    Dim IsLost As Boolean
    Dim IsFirstGS As Boolean
    Dim IsInGS As Boolean
    Dim IsInFoodPatch As Boolean
    Dim IsUpdateActivity As Boolean
    Dim WasLost As Boolean
        
    ' Miscellaneous variables
    Dim CurrentWS As Worksheet
    Dim CurrentSpecies As String

    ' Initialize the array of worksheet names
    Call Read_WorksheetNames(WorksheetNames)
   
    ' Read in the behavior codes, ID codes, and enumeration text
    Call Read_Codes_All
    
    ' Find last row of input
    lastRow = Get_LastRow
    
    ' Make sure that there is data to parse
    If lastRow = 0 Then
        MsgBox ("Add some data first.")
        Exit Sub
    End If
   
    ' Read the input data and store it in the InputData array
    InputData = Read_InputData(lastRow)
    
    ' Check that all the input data are correct
    ' Exit program if there are input errors
    If Not IsInputOK(InputData) Then
        'MsgBox ("Correct Data Entry Errors Before Continuing")
        Exit Sub
    End If

    ' The active worksheet will be the one from which ProcessData was called
    ' It is assumed that this has the focal data in a correct format
    Set CurrentWS = ActiveSheet

    ' Add (and/or Clear) the output worksheets
    Call Worksheets_Add(WorksheetNames)
    ' Initialize the worksheets and write column names
    Call Write_Headers
   
    ' Be sure the data worksheet is the active worksheet
    CurrentWS.Activate
    
    With Cells.Font
        .name = "Consolas"
        .Size = 10
    End With
    
    ReDim VertebrateData(0)
    ReDim TreeCBHData(0)
    CurrentGPSColor = ""
    IsFirstGS = True
    IsInGS = False
    IsLost = False
    IsInFollow = False
    IsInFoodPatch = False
    IsUpdateActivity = False
    WasLost = False
    
    ' Process data line by line until end of InputData
    For RowIn = 1 To UBound(InputData)
        Select Case InputData(RowIn).LineType
            
            ' Ignore these and move along
            Case LineTypeEnum.Blank
            Case LineTypeEnum.Other
                                                    
            ' Header and DONE lines
            Case LineTypeEnum.HeaderObserver
                Call Update_Observation(CurrentObservation, InputData, RowIn, WasLost)
                CurrentFollow.SeqNum = 0
                IsLost = False
            Case LineTypeEnum.HeaderGroup
                Call Update_Observation(CurrentObservation, InputData, RowIn, WasLost)
                IsLost = False
            Case LineTypeEnum.GPSColor
                CurrentGPSColor = Mid(InputData(RowIn).Data, 5)
            Case LineTypeEnum.Done
                If CurrentObservation.FindType = "Wake" And CurrentObservation.LeaveType = "Sleep" Then
                    CurrentObservation.IsFullDay = True
                Else
                    CurrentObservation.IsFullDay = False
                End If
                Call Write_Observation(CurrentObservation)
                
                If IsInGS Then
                    Call Write_GroupScan(CurrentGroupScan)
                    IsInGS = False
                End If
                
            ' Group scan lines
            Case LineTypeEnum.GSWaypoint
                
                ' Enters here if NOT first group scan and year is 2010 or 2009
                If Not IsFirstGS And Year(InputData(RowIn).Datim) <> 2011 Then
                    Call Write_GroupScan(CurrentGroupScan)
                End If
                
                CurrentVertebrate.VertSeqNum = 0
                CurrentGroupScan.Climate = ""
                CurrentGroupScan.CanopyHeight = -1
                CurrentGroupScan.ForestLevel = ""
                CurrentGroupScan.GroupActivity = ""
                CurrentGroupScan.GroupHeight = -1
                CurrentGroupScan.SpeciesCode = ""
                CurrentGroupScan.Stage = -1
                Call Update_GroupScan(CurrentGroupScan, CurrentObservation, InputData, RowIn)
                IsFirstGS = False
                IsInGS = True
            Case LineTypeEnum.GSClimate, LineTypeEnum.GSHeight, LineTypeEnum.GSLevel, LineTypeEnum.GSStage, LineTypeEnum.GSActivity
                Call Update_GroupScan(CurrentGroupScan, CurrentObservation, InputData, RowIn)
            Case LineTypeEnum.GSVertebrate
'                CurrentVertResponse.VertResponseSeqNum = 0
                If Mid(InputData(RowIn).Data, 4) <> "0" Then
                    Call Update_Vertebrate(CurrentVertebrate, VertebrateData, CurrentGroupScan, NextNewVertebrateID, InputData, RowIn)
                End If
'            Case LineTypeEnum.ScanResponse
'                If Mid(InputData(RowIn).Data, 4) <> "0" Then
'                    Call Update_VertResponse(CurrentVertResponse, VertResponseData, CurrentVertebrate, NextNewVertResponseID, InputData, RowIn)
'                End If
            Case LineTypeEnum.GSEnd
                Call Write_GroupScan(CurrentGroupScan)
                If UBound(VertebrateData) <> 0 Then
                    For i = 1 To UBound(VertebrateData)
                        Call Write_Vertebrate(VertebrateData(i))
                    Next i
                End If
'                If UBound(VertResponseData) <> 0 Then
'                    For i = 1 To UBound(VertResponseData)
'                        Call Write_VertResponse(VertResponseData(i))
'                    Next i
'                End If
                ReDim VertebrateData(0)
'                ReDim VertResponseData(0)
                
                IsInGS = False
                
                                        
            ' Tree lines
            Case LineTypeEnum.TreeID
                CurrentTreeCBH.StemNum = 0
                Call Update_FruitVisit(CurrentFruitVisit, CurrentObservation, InputData, RowIn)
            Case LineTypeEnum.TreeNum, LineTypeEnum.TreeSpecies, LineTypeEnum.TreeWaypoint, LineTypeEnum.TreePhenology, LineTypeEnum.TreeDisks, LineTypeEnum.TreeBromeliads
                Call Update_FruitVisit(CurrentFruitVisit, CurrentObservation, InputData, RowIn)
            Case LineTypeEnum.TreeCBH
                TreeCBHData = Parse_TreeCBH(InputData(RowIn).Data)
                For i = 1 To UBound(TreeCBHData)
                    Call Update_TreeCBH(CurrentTreeCBH, CurrentFruitVisit, TreeCBHData(i))
                    Call Write_TreeCBH(CurrentTreeCBH)
                Next i
                ReDim TreeCBHData(0)
            Case LineTypeEnum.TreeEnd
                If CurrentFruitVisit.SpeciesCode = "" Then CurrentFruitVisit.SpeciesCode = "USPE"
                Call Write_FruitVisit(CurrentFruitVisit)
            
            ' Ranging lines
            Case LineTypeEnum.RangingSleep, LineTypeEnum.RangingWake
                Call Update_Observation(CurrentObservation, InputData, RowIn, WasLost)
            Case LineTypeEnum.RangingFind
                Call Update_Observation(CurrentObservation, InputData, RowIn, WasLost)
            Case LineTypeEnum.RangingLeave
                Call Update_Observation(CurrentObservation, InputData, RowIn, WasLost)
                Call Write_Observation(CurrentObservation)
            Case LineTypeEnum.RangingWater
                Call Update_RangingEvent(CurrentRangingEvent, CurrentObservation, InputData, RowIn)
                Call Write_RangingEvent(CurrentRangingEvent)
            Case LineTypeEnum.RangingVertebrate
                Call Update_RangingEvent(CurrentRangingEvent, CurrentObservation, InputData, RowIn)
                Call Write_RangingEvent(CurrentRangingEvent)
                
            ' Follow lines
            Case LineTypeEnum.FollowStart
                ' Reset the flag for being in a follow
                IsInFollow = True
                
                ' Update current follow and write one record to follow table
                Call Update_Follow(CurrentFollow, CurrentObservation, InputData, RowIn)
                Call Write_Follow(CurrentFollow)
                
                ' Update current follow block and write one record to follow block table
                If CurrentFollow.FollowType = "Normal" Then
                    Call Update_FollowBlock(CurrentFollowBlock, CurrentFollow, InputData, RowIn)
                    If CurrentFollowBlock.DurationOfBlock <> 0 Then Call Write_FollowBlock(CurrentFollowBlock)
                End If
                
                ' A new follow resets the PointSample, SelfDirected, and ForagingEvent sequence numbers
                CurrentPointSample.SeqNum = 0
                CurrentSelfDirected.SeqNum = 0
                CurrentForagingEvent.SeqNum = 0
            
            Case LineTypeEnum.Abort, LineTypeEnum.FollowEnd
                IsInFollow = False
                CurrentSpecies = ""
                If IsInFoodPatch Then
                    Call Write_FoodPatch(CurrentFoodPatch)
                    IsInFoodPatch = False
                End If
                
            Case LineTypeEnum.FollowBlockStatus
                If InputData(RowIn).Data = "II" And CurrentFollowBlock.IsInActivity = False Then IsUpdateActivity = True
                
                Call Update_FollowBlock(CurrentFollowBlock, CurrentFollow, InputData, RowIn)
                If CurrentFollowBlock.DurationOfBlock <> 0 Then Call Write_FollowBlock(CurrentFollowBlock)
                
                ' Update Activity for IA & II
                If InputData(RowIn).Data = "IA" Or IsUpdateActivity Then
                    Call Update_Activity(CurrentActivity, CurrentFollow, CurrentFollowBlock, InputData, RowIn)
                    If CurrentActivity.DurationOfState <> 0 Then Call Write_Activity(CurrentActivity)
                End If
                IsUpdateActivity = False
                

            ' Point sample lines
            Case LineTypeEnum.PSActivity
                Call Update_PointSample(CurrentPointSample, CurrentFollow, CurrentFollowBlock, CurrentSpecies, InputData, RowIn)
                Call Write_PointSample(CurrentPointSample)
                    If CurrentPointSample.SeqNum = 1 Or CurrentActivity.Activity = "O" Then
                        Call Update_Activity(CurrentActivity, CurrentFollow, CurrentFollowBlock, InputData, RowIn)
                        If CurrentActivity.DurationOfState <> 0 Then Call Write_Activity(CurrentActivity)
                    End If
            
            ' Follow states and events
            Case LineTypeEnum.FollowActivity
                Call Update_Activity(CurrentActivity, CurrentFollow, CurrentFollowBlock, InputData, RowIn)
                If CurrentActivity.DurationOfState <> 0 Then Call Write_Activity(CurrentActivity)
    
            Case LineTypeEnum.SelfDirected
                Call Update_SelfDirected(CurrentSelfDirected, CurrentFollow, CurrentFollowBlock, CurrentActivity, InputData, RowIn)
                Call Write_SelfDirected(CurrentSelfDirected)
    
            ' Foraging lines
            Case LineTypeEnum.EatNew
                Call Update_FoodObject(CurrentFoodObject, CurrentFollow, CurrentFollowBlock, CurrentFoodPatch, CurrentSpecies, IsInFoodPatch, InputData, RowIn)
                Call Write_FoodObject(CurrentFoodObject)
                CurrentForagingEvent.SeqNum = 0
                Call Update_ForagingEvent(CurrentForagingEvent, CurrentFollow, CurrentFoodObject, InputData, RowIn)
                Call Write_ForagingEvent(CurrentForagingEvent)
            Case LineTypeEnum.EatSame
                Call Update_ForagingEvent(CurrentForagingEvent, CurrentFollow, CurrentFoodObject, InputData, RowIn)
                Call Write_ForagingEvent(CurrentForagingEvent)
    
            'Food patch lines
            Case LineTypeEnum.FoodPatchEnter
                Call Update_FoodPatch(CurrentFoodPatch, CurrentFollow, CurrentFollowBlock, CurrentSpecies, InputData, RowIn)
                IsInFoodPatch = True
            Case LineTypeEnum.FoodPatchEnd
                Call Write_FoodPatch(CurrentFoodPatch)
                IsInFoodPatch = False
                
            ' Alarm lines
            Case LineTypeEnum.Alarm, LineTypeEnum.AlarmPresent, LineTypeEnum.AlarmDanger, LineTypeEnum.AlarmIntensity, LineTypeEnum.AlarmLevel, LineTypeEnum.AlarmMultiple, LineTypeEnum.AlarmWaypoint, LineTypeEnum.AlarmSpecies
                Call Update_Alarm(CurrentAlarm, CurrentObservation, InputData, RowIn)
            Case LineTypeEnum.AlarmEnd
                Call Write_Alarm(CurrentAlarm)
    
            ' Behavior lines
            Case LineTypeEnum.Behavior
                Call Update_Interaction(CurrentInteraction, CurrentObservation, InputData, RowIn)
                Call Write_Interaction(CurrentInteraction)
            Case LineTypeEnum.BehaviorIntergroup
                Call Update_Intergroup(CurrentIntergroup, CurrentObservation, InputData, RowIn)
                Call Write_Intergroup(CurrentIntergroup)
    
            ' Other lines
            Case LineTypeEnum.Comment
                Call Update_Comment(CurrentComment, CurrentObservation, InputData, RowIn)
                Call Write_Comment(CurrentComment)
                
        End Select

    Next RowIn

Finished:
    ' If there are no errors, display all the output worksheets
    Call Worksheets_Reveal(WorksheetNames)
    CurrentWS.Activate

    ' Put the focus to the top left cell
    Range("A1").Select

    ' Sound a tone to indicate that the program is finished
    Beep
    Exit Sub

Errors:
    MsgBox ("Fatal error on line " & RowIn & vbCrLf & "Psion Input = " & InputData(RowIn).Data & vbCrLf & "Correct the error and reparse the data")
    GoTo Finished
End Sub

' This sub calls a function to add blank worksheets
' If one or more already exist, it clears whatever was previously there
' They remain hidden in case there is an error until Worksheets_Reveal is called
Private Sub Worksheets_Add(ByRef strWorksheetNames() As String)
    Application.DisplayAlerts = False
    Dim i As Integer
    
    For i = LBound(strWorksheetNames) To UBound(strWorksheetNames)
        Call WorkSheet_Add(strWorksheetNames(i))
    Next i
    
    Call Worksheets_Hide(strWorksheetNames)
    Call Worksheets_Clear(strWorksheetNames)
    
    Application.DisplayAlerts = True
End Sub

' This sub adds one blank worksheet with whatever name is passed to it
Private Sub WorkSheet_Add(strSheetName As String)
    Dim intI As Integer
    Dim intCount As Integer

    intCount = Worksheets.Count
    For intI = 1 To intCount
        If Sheets(intI).name = strSheetName Then Exit Sub
    Next intI

    Worksheets.Add Count:=1, After:=Sheets(intCount)
    Sheets(intCount + 1).name = strSheetName
    
    With Cells.Font
        .name = "Consolas"
        .Size = 10
    End With

End Sub

' This sub is called when the program finishes running correctly
' It makes the 16 output worksheets visible
Private Sub Worksheets_Reveal(ByRef strSheetNames() As String)
    Application.DisplayAlerts = False
    Dim i As Integer
    
    For i = LBound(strSheetNames) To UBound(strSheetNames)
        Sheets(strSheetNames(i)).Visible = True
    Next i
    
    Application.DisplayAlerts = True

End Sub

' Makes the worksheet not visible
Private Sub Worksheets_Hide(ByRef strSheetNames() As String)
    Application.DisplayAlerts = False
    Dim i As Integer
    For i = LBound(strSheetNames) To UBound(strSheetNames)
        Sheets(strSheetNames(i)).Visible = False
    Next i
    Application.DisplayAlerts = True
End Sub

' Calls a function to clear the worksheets
Private Sub Worksheets_Clear(ByRef strSheetNames() As String)
    Dim i As Integer
    For i = LBound(strSheetNames) To UBound(strSheetNames)
        Call Worksheet_Clear(strSheetNames(i))
    Next i
End Sub

' Clears a specific worksheet
Private Sub Worksheet_Clear(ByVal strName As String)
    Dim i As Integer
    For i = 1 To Sheets.Count
        ' If the worksheet already exits, activate it and clear it, and then exit
        If Sheets(i).name = strName Then
            Worksheets(i).Activate
            Cells.Select
            Selection.Delete Shift:=xlUp
            Range("A1").Select
            Exit Sub
        End If
    Next i
End Sub

' Create a specific worksheet
Private Sub Worksheet_Create(ByVal strName As String)
    Dim i As Integer
    For i = 1 To Sheets.Count
        ' If the worksheet already exits, activate it and clear it, and then exit
        If Sheets(i).name = strName Then
            Call Worksheet_Clear(strName)
            Exit Sub
        End If
    Next i

    ' If the worksheet doesn't exist, then just create it and activate it
    Worksheets.Add Count:=1, After:=Sheets(Sheets.Count)
    Worksheets(Sheets.Count).Activate
    ActiveSheet.name = strName
End Sub

' Delete a specific worksheet
Public Sub Worksheet_Delete(ByVal strName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(strName).Delete
    Application.DisplayAlerts = True
End Sub

' This function checks all input lines for correct syntax
' It returns true if there are no input errors and false if input errors exist
' In the case of input errors, it calls subs to write them to the input errors worksheet
Private Function IsInputOK(ByRef InputData() As InputLine) As Boolean

    Dim RetCode As Long
    Dim i As Long
    Dim j As Long
    Dim NInputErrors As Integer
    Dim IsInGroupScan As Boolean
    Dim IsInFollow As Boolean
    Dim IsInFruitVisit As Boolean
    Dim IsInFoodPatch As Boolean
    Dim IsInAlarm As Boolean
    Dim IsSetZ1 As Boolean
    Dim IsSetZ2 As Boolean
    Dim IsSetSub As Boolean
    Dim IsSetPos As Boolean
    Dim IsSetLev As Boolean
    Dim IsSetSpecies As Boolean
    Dim IsNormalFollow As Boolean
    Dim TestVisStatus As Boolean
    
    Dim OpenFoodObjects(15) As Boolean
    Dim VisibilityStatus(4) As Boolean
    
    IsInGroupScan = False
    IsInFollow = False
    IsInFruitVisit = False
    IsInFoodPatch = False
    IsInAlarm = False
    IsInputOK = True
    IsSetZ1 = False
    IsSetZ2 = False
    IsSetSub = False
    IsSetPos = False
    IsSetLev = False
    IsSetSpecies = False
    IsNormalFollow = False
    TestVisStatus = False
    
    NInputErrors = 0
    
    ' Read each line from the input range(1 to LastRow) and check for errors
    For i = 1 To UBound(InputData)
        
        ' Check to make sure day begins with header lines
        If i = 1 Then
            If InputData(i).LineType <> LineTypeEnum.HeaderObserver Then
                If NInputErrors = 0 Then
                    ' create/activate the errors worksheet
                    Call Write_InputError_Header
                End If
                NInputErrors = NInputErrors + 1
                Call Write_InputError(NInputErrors, InputData(i - 1), "Day should begin with observer code")
                IsInputOK = False
            End If
        End If
        If i = 2 Then
            If InputData(i).LineType <> LineTypeEnum.HeaderGroup Then
                If NInputErrors = 0 Then
                    ' create/activate the errors worksheet
                    Call Write_InputError_Header
                End If
                NInputErrors = NInputErrors + 1
                Call Write_InputError(NInputErrors, InputData(i - 1), "Second line should be group code")
                IsInputOK = False
            End If
        End If
        
        ' Select line type for the current line
        Select Case InputData(i).LineType
            
            ' Handle code errors
            Case LineTypeEnum.Unknown
                ' Output an error for any unknown line types
                If NInputErrors = 0 Then
                    ' Create/activate the errors worksheet
                    Call Write_InputError_Header
                End If
                NInputErrors = NInputErrors + 1
                Call Write_InputError(NInputErrors, InputData(i), "Code error")
                IsInputOK = False
            
            Case LineTypeEnum.Done
                If Hour(InputData(i).Datim) > 18 Or (Hour(InputData(i).Datim) = 18 And Minute(InputData(i).Datim) >= 50) Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "DONE statement is too late. Put at end of monkey contact period.")
                    IsInputOK = False
                End If
                If IsInFruitVisit Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Day ends with missing tree end code")
                    IsInputOK = False
                End If
                If IsInGroupScan And Year(InputData(i).Datim) = 2011 Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Day ends with missing group scan end code")
                    IsInputOK = False
                End If
                If IsInAlarm Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Day ends with missing alarm end code")
                    IsInputOK = False
                End If
                If IsInFoodPatch Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Day ends with missing CX code")
                    IsInputOK = False
                End If
                If IsInFollow Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Day ends with missing END code")
                    IsInputOK = False
                End If
            
            ' Check tree codes (begins with TI and ends with TX)
            Case LineTypeEnum.TreeID
                If IsInFruitVisit Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Tree visit begins before previous one ends")
                    IsInputOK = False
                End If
                IsInFruitVisit = True
            Case LineTypeEnum.TreeNum, LineTypeEnum.TreeCBH, LineTypeEnum.TreePhenology, LineTypeEnum.TreeSpecies, LineTypeEnum.TreeWaypoint, LineTypeEnum.TreeEnd
                If Not IsInFruitVisit Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Tree visit does not begin with TI")
                    IsInputOK = False
                End If
                If InputData(i).LineType = LineTypeEnum.TreeEnd Then IsInFruitVisit = False
                
                
            ' Check group scan codes (begins with GW and ends with GX)
            Case LineTypeEnum.GSWaypoint
                If IsInGroupScan And Year(InputData(i).Datim) = 2011 Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Group scan begins before previous one ends")
                    IsInputOK = False
                End If
                IsInGroupScan = True
            Case LineTypeEnum.GSActivity, LineTypeEnum.GSClimate, LineTypeEnum.GSEnd, LineTypeEnum.GSHeight, LineTypeEnum.GSLevel, LineTypeEnum.GSStage, LineTypeEnum.GSVertebrate
                If Not IsInGroupScan Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Group scan does not begin with GW")
                    IsInputOK = False
                End If
                If InputData(i).LineType = LineTypeEnum.GSEnd Then IsInGroupScan = False
                
            ' Check alarm codes (begins with VS, VB, VU, or VT and ends with VX)
            Case LineTypeEnum.Alarm
                If IsInAlarm Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Alarm begins before previous one ends")
                    IsInputOK = False
                End If
                IsInAlarm = True
            Case LineTypeEnum.AlarmDanger, LineTypeEnum.AlarmEnd, LineTypeEnum.AlarmIntensity, LineTypeEnum.AlarmLevel, LineTypeEnum.AlarmMultiple, LineTypeEnum.AlarmPresent, LineTypeEnum.AlarmSpecies, LineTypeEnum.AlarmWaypoint
                If Not IsInAlarm Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Alarm codes do not beging with VS, VB, VT, or VU")
                    IsInputOK = False
                End If
                If InputData(i).LineType = LineTypeEnum.AlarmEnd Then IsInAlarm = False
                           
            ' Check follow codes (begins with F:__ and ends with END or ABORT)
            Case LineTypeEnum.FollowStart
                If IsInFollow Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "No END code for previous follow")
                    IsInputOK = False
                End If
                IsInFollow = True
                
                For j = 1 To 15
                    OpenFoodObjects(j) = False
                Next j
                For j = 1 To 4
                    VisibilityStatus(j) = True
                Next j
                
                If Left(InputData(i).Data, 2) = "F:" Or Left(InputData(i).Data, 2) = "F." Then
                    IsNormalFollow = True
                End If
                
            Case LineTypeEnum.RangingWake
                If Hour(InputData(i).Datim) > 5 Or Hour(InputData(i).Datim) < 3 Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Unrealistic wake time")
                    IsInputOK = False
                End If
            
            Case LineTypeEnum.RangingSleep
                If Hour(InputData(i).Datim) < 17 Or Hour(InputData(i).Datim) > 18 Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Unrealistic sleep time")
                    IsInputOK = False
                End If
            
            Case LineTypeEnum.EatNew, LineTypeEnum.EatSame, LineTypeEnum.FollowActivity, LineTypeEnum.FollowStatus, _
                LineTypeEnum.FoodPatchEnter, LineTypeEnum.FoodPatchEnd, LineTypeEnum.GPSError, LineTypeEnum.PSActivity, _
                LineTypeEnum.PSLevel, LineTypeEnum.PSPosture, LineTypeEnum.PSSubstrate, LineTypeEnum.SelfDirected, _
                LineTypeEnum.FollowNoMovement, LineTypeEnum.FollowEnd, LineTypeEnum.Abort, LineTypeEnum.FollowBlockStatus
                If Not IsInFollow Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i), "Follow code occurs outside of follow")
                    IsInputOK = False
                End If
                
                Select Case InputData(i).LineType
                    Case LineTypeEnum.FollowEnd, LineTypeEnum.Abort
                        Select Case InputData(i).LineType
                            Case LineTypeEnum.FollowEnd
                                Select Case Left(InputData(i).Data, 2)
                                    Case "EN"
                                        If Not IsNormalFollow Then
                                            If NInputErrors = 0 Then
                                                ' create/activate the errors worksheet
                                                Call Write_InputError_Header
                                            End If
                                            NInputErrors = NInputErrors + 1
                                            Call Write_InputError(NInputErrors, InputData(i), "Scan or feeding follow should end with FX")
                                            IsInputOK = False
                                        End If
                                    Case "FX"
                                        If IsNormalFollow Then
                                            If NInputErrors = 0 Then
                                                ' create/activate the errors worksheet
                                                Call Write_InputError_Header
                                            End If
                                            NInputErrors = NInputErrors + 1
                                            Call Write_InputError(NInputErrors, InputData(i), "Normal follow ends with FX")
                                            IsInputOK = False
                                        End If
                                End Select
                            Case LineTypeEnum.Abort
                                If Not IsNormalFollow Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Scan or feeding follow ends with ABORT")
                                    IsInputOK = False
                                End If
                        End Select
                        IsInFollow = False
                        IsInFoodPatch = False
                        IsSetSpecies = False
                        IsSetZ1 = False
                        IsSetZ2 = False
                        IsNormalFollow = False
                    
                    Case LineTypeEnum.PSActivity
                        If Not VisibilityStatus(VisibilityEnum.V_PointSample) Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Point sample during OP block")
                            IsInputOK = False
                        End If
                        If InStr(InputData(i).Data, ".") And Not IsSetSpecies And Not Has_SpeciesCode(InputData(i).Data) Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Species not set for . following PS activity")
                            IsInputOK = False
                        End If
                        If Has_SpeciesCode(InputData(i).Data) Then IsSetSpecies = True
                        Select Case Mid(InputData(i).Data, 2, 1)
                            Case "F", "E", "X"
                                Select Case Get_FoodObjectIndex(Mid(InputData(i).Data, 3, 1))
                                    Case FoodObjectEnum.Flower, FoodObjectEnum.Fruit, FoodObjectEnum.Leaf, FoodObjectEnum.Pith, FoodObjectEnum.Seed, FoodObjectEnum.Thorn
                                        If IsSetSpecies And InStr(InputData(i).Data, ".") = 0 Then
                                            If NInputErrors = 0 Then
                                                ' create/activate the errors worksheet
                                                Call Write_InputError_Header
                                            End If
                                            NInputErrors = NInputErrors + 1
                                            Call Write_InputError(NInputErrors, InputData(i), "Check PS for missing .")
                                            IsInputOK = False
                                        End If
                                End Select
                        End Select
                        IsSetSub = False
                        IsSetLev = False
                        IsSetPos = False
                        
                    Case LineTypeEnum.PSLevel
                        If IsSetLev Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Duplicate PS level code")
                            IsInputOK = False
                        End If
                        IsSetLev = True
                                        
                    Case LineTypeEnum.PSSubstrate
                        If IsSetSub Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Duplicate PS substrate code")
                            IsInputOK = False
                        End If
                        IsSetSub = True
                        
                    Case LineTypeEnum.PSPosture
                        If IsSetPos Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Duplicate PS posture code")
                            IsInputOK = False
                        End If
                        IsSetPos = True
                    
                    Case LineTypeEnum.EatNew
                        If Not VisibilityStatus(VisibilityEnum.V_Forage) Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Eat code during OF block")
                            IsInputOK = False
                        End If
                        OpenFoodObjects(Get_FoodObjectIndex(Mid(InputData(i).Data, 2, 1))) = True
                        If InStr(InputData(i).Data, ".") And Not IsSetSpecies And Not Has_SpeciesCode(InputData(i).Data) Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Species not set for . following eat code")
                            IsInputOK = False
                        End If
                        If Has_SpeciesCode(InputData(i).Data) Then IsSetSpecies = True
                    
                    Case LineTypeEnum.EatSame
                        If Not VisibilityStatus(VisibilityEnum.V_Forage) Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Eat code during OF block")
                            IsInputOK = False
                        End If
                        If Not OpenFoodObjects(Get_FoodObjectIndex(Mid(InputData(i).Data, 2, 1))) Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Eat same without preceeding eat new")
                            IsInputOK = False
                        End If
                        
                    ' Check food patch codes (begins with C_ and ends with CX)
                    Case LineTypeEnum.FoodPatchEnter
                        If IsInFoodPatch Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Missing CX code")
                            IsInputOK = False
                        End If
                        IsInFoodPatch = True
                        If Has_SpeciesCode(InputData(i).Data) Then IsSetSpecies = True
                    Case LineTypeEnum.FoodPatchEnd
                        If Not IsInFoodPatch Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Missing food patch begin code")
                            IsInputOK = False
                        End If
                        IsInFoodPatch = False
                    
                    Case LineTypeEnum.GPSError
                        Select Case Mid(InputData(i).Data, 2, 1)
                            Case "1"
                                If IsSetZ1 Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Z1 already set")
                                    IsInputOK = False
                                Else
                                    IsSetZ1 = True
                                End If
                            Case "2"
                                If IsSetZ2 Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Z2 already set")
                                    IsInputOK = False
                                Else
                                    IsSetZ2 = True
                                End If
                        End Select
                    
                    Case LineTypeEnum.FollowActivity
                        If Not VisibilityStatus(VisibilityEnum.V_Activity) Then
                            If NInputErrors = 0 Then
                                ' create/activate the errors worksheet
                                Call Write_InputError_Header
                            End If
                            NInputErrors = NInputErrors + 1
                            Call Write_InputError(NInputErrors, InputData(i), "Activity change during OA block")
                            IsInputOK = False
                        End If
                        
                    Case LineTypeEnum.FollowBlockStatus
                        Select Case InputData(i).Data
                            Case "OO"
                                TestVisStatus = False
                                For j = 1 To 4
                                    If VisibilityStatus(j) Then
                                       TestVisStatus = True
                                    End If
                                    VisibilityStatus(j) = False
                                Next j
                                If Not TestVisStatus Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant OO")
                                    IsInputOK = False
                                End If
                            Case "OF"
                                If Not VisibilityStatus(VisibilityEnum.V_Forage) Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant OF")
                                    IsInputOK = False
                                End If
                                VisibilityStatus(VisibilityEnum.V_Forage) = False
                            Case "OA"
                                If Not VisibilityStatus(VisibilityEnum.V_Activity) Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant OA")
                                    IsInputOK = False
                                End If
                                VisibilityStatus(VisibilityEnum.V_Activity) = False
                            Case "OP"
                                If Not VisibilityStatus(VisibilityEnum.V_PointSample) Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant OP")
                                    IsInputOK = False
                                End If
                                VisibilityStatus(VisibilityEnum.V_PointSample) = False
                            Case "OT"
                                If Not VisibilityStatus(VisibilityEnum.V_Track) Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant OT")
                                    IsInputOK = False
                                End If
                                VisibilityStatus(VisibilityEnum.V_Track) = False
                            Case "II"
                                TestVisStatus = True
                                For j = 1 To 4
                                    If Not VisibilityStatus(j) Then
                                        TestVisStatus = False
                                    End If
                                    VisibilityStatus(j) = True
                                Next j
                                If TestVisStatus Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant II")
                                    IsInputOK = False
                                End If
                            Case "IA"
                                If VisibilityStatus(VisibilityEnum.V_Activity) Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant IA")
                                    IsInputOK = False
                                End If
                                VisibilityStatus(VisibilityEnum.V_Activity) = True
                            Case "IF"
                                If VisibilityStatus(VisibilityEnum.V_Forage) Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant IF")
                                    IsInputOK = False
                                End If
                                VisibilityStatus(VisibilityEnum.V_Forage) = True
                            Case "IP"
                                If VisibilityStatus(VisibilityEnum.V_PointSample) Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant IP")
                                    IsInputOK = False
                                End If
                                VisibilityStatus(VisibilityEnum.V_PointSample) = True
                            Case "IT"
                                If VisibilityStatus(VisibilityEnum.V_Track) Then
                                    If NInputErrors = 0 Then
                                        ' create/activate the errors worksheet
                                        Call Write_InputError_Header
                                    End If
                                    NInputErrors = NInputErrors + 1
                                    Call Write_InputError(NInputErrors, InputData(i), "Redundant IT")
                                    IsInputOK = False
                                End If
                                VisibilityStatus(VisibilityEnum.V_Track) = True
                        End Select
                    
                End Select
        End Select
    
        ' Check for proper sequence of times
        If i >= 2 Then
            If InputData(i).Datim < InputData(i - 1).Datim Then
                If NInputErrors = 0 Then
                    ' create/activate the errors worksheet
                    Call Write_InputError_Header
                End If
                NInputErrors = NInputErrors + 1
                Call Write_InputError(NInputErrors, InputData(i), "Date/Time is out of sequence")
                IsInputOK = False
            End If
            
            ' Check to make sure day ends with "DONE"
            If i = UBound(InputData) Then
                If InputData(i).Data <> "DONE" Then
                    If NInputErrors = 0 Then
                        ' create/activate the errors worksheet
                        Call Write_InputError_Header
                    End If
                    NInputErrors = NInputErrors + 1
                    Call Write_InputError(NInputErrors, InputData(i - 1), "Day ends without DONE")
                    IsInputOK = False
                End If
            End If
            
            ' Make sure dates are all the same
            If DateValue(InputData(1).Datim) <> DateValue(InputData(i).Datim) Then
                If NInputErrors = 0 Then
                    ' create/activate the errors worksheet
                    Call Write_InputError_Header
                End If
                NInputErrors = NInputErrors + 1
                Call Write_InputError(NInputErrors, InputData(i - 1), "Date change")
                IsInputOK = False
            End If
            
        End If
        
    Next i
    
    If NInputErrors > 0 Then
        RetCode = MsgBox("There were " & NInputErrors & " input lines that could not be parsed", , "Fatal Error")
    Else
        ' Delete the errors worksheet if it exists
        Call Worksheet_Delete("InputError")
    End If
    Exit Function
End Function

' This function returns the last used row of the active worksheet
Private Function Get_LastRow() As Long
    Dim lastRow As Long
    If WorksheetFunction.CountA(Cells) > 0 Then
        ' Search for any entry, by searching backwards by rows
        lastRow = Cells.Find(What:="*", After:=[A1], _
              SearchOrder:=xlByRows, _
              SearchDirection:=xlPrevious).Row
              Get_LastRow = lastRow
    End If
End Function

Private Function StringIsInteger(testString As String) As Boolean

    Dim asStr As String
    Dim tempStr As String
    
    tempStr = Trim(testString)
    StringIsInteger = False
    
    'Trim leading + or -
    If (InStr(tempStr, "+") = 1 Or InStr(tempStr, "-") = 1) Then
        tempStr = Mid(tempStr, 2)
    End If
    
    ' Trim multiple leading zeros (leaves one if number is '0')
    While (Len(tempStr) > 1 And InStr(tempStr, "0") = 1)
        tempStr = Mid(tempStr, 2)
    Wend
    
    If IsNumeric(tempStr) Then
        asStr = CStr(CLng(tempStr))
    
        If (tempStr = asStr) Then
            StringIsInteger = True
        End If
    
    End If

End Function

' This function reads and stores all input lines in an array of InputLine types
' Column A must contain data/time information
' Column B must contain recognizable data codes
Private Function Read_InputData(ByVal lastRow As Long) As InputLine()
    Dim InputData() As InputLine
    Dim inputRow As Long
   
    ' Read each line of input from the input range (1 to last row)
    ReDim InputData(1 To lastRow)
   
    ' Trim leading blanks, convert to all caps, read the first four columns
    For inputRow = 1 To lastRow
        
        ' Get date/time from column 1
        InputData(inputRow).Datim = Cells(inputRow, 1)
        
        ' Get Psion data from column 2
        InputData(inputRow).Data = UCase(Trim(Cells(inputRow, 2)))
        
        ' Determine the line type by passing the Psion data to the Get_LineType function
        InputData(inputRow).LineType = Get_LineType(InputData(inputRow).Data)
        
        ' Store input row number
        InputData(inputRow).LineNum = inputRow
        
    Next inputRow
    
    ' Return the InputData array
    Read_InputData = InputData
    
End Function

Public Sub Read_WorksheetNames(ByRef WorksheetNames() As String)
    
    ReDim WorksheetNames(0 To 17)

    ' Set the names for all output worksheets
    WorksheetNames(0) = "Observation"
    WorksheetNames(1) = "GroupScan"
    WorksheetNames(2) = "Vertebrate"
    WorksheetNames(3) = "Follow"
    WorksheetNames(4) = "FollowBlock"
    WorksheetNames(5) = "PointSample"
    WorksheetNames(6) = "Activity"
    WorksheetNames(7) = "SelfDirected"
    WorksheetNames(8) = "FoodPatch"
    WorksheetNames(9) = "FoodObject"
    WorksheetNames(10) = "ForagingEvent"
    WorksheetNames(11) = "FruitVisit"
    WorksheetNames(12) = "TreeCBH"
    WorksheetNames(13) = "Alarm"
    WorksheetNames(14) = "Interaction"
    WorksheetNames(15) = "Intergroup"
    WorksheetNames(16) = "RangingEvent"
    WorksheetNames(17) = "Comment"
    
End Sub

' This sub calls the function Read_Codes for each set of codes on the Codes worksheet
Private Sub Read_Codes_All()

    Call Read_Codes("ObserverCodes", ObserverCodes)
    Call Read_Codes("GroupCodes", GroupCodes)
    Call Read_Codes("TreeCodes", TreeCodes)
    Call Read_Codes("FoodCodes", FoodCodes)
    Call Read_Codes("GroupScanCodes", GroupScanCodes)
    Call Read_Codes("HeaderCodes", HeaderCodes)
    Call Read_Codes("RangingCodes", RangingCodes)
    Call Read_Codes("ClimateCodes", ClimateCodes)
    Call Read_Codes("VertebrateCodes", VertebrateCodes)
    Call Read_Codes("PointSampleCodes", PointSampleCodes)
    Call Read_Codes("LevelCodes", LevelCodes)
    Call Read_Codes("ActivityCodes", ActivityCodes)
    Call Read_Codes("FoodPatchCodes", FoodPatchCodes)
    Call Read_Codes("SelfDirectedCodes", SelfDirectedCodes)
    Call Read_Codes("EatCodes", EatCodes)
    Call Read_Codes("PostureCodes", PostureCodes)
    Call Read_Codes("SubstrateCodes", SubstrateCodes)
    Call Read_Codes("AlarmCodes", AlarmCodes)
    Call Read_Codes("FollowStatusCodes", FollowStatusCodes)
    Call Read_Codes("FollowPartCodes", FollowPartCodes)
    Call Read_Codes("MonkeyCodes", MonkeyCodes)
    Call Read_Codes("GPSUnits", GPSUnits)
    Call Read_Codes("CentralityCodes", CentralityCodes)
    
End Sub

' This sub reads in and stores values from the specified set of codes on the Codes worksheet
Private Sub Read_Codes(ByVal NameOfCodes As String, ByRef Codes() As String)
    Dim intRowLast, i As Integer
    Dim Value As String
    Dim MyRange As Range
    
    ' throw an error if NameOfCodes is not valid
    On Error GoTo BadRangeName
    
    ' Create a range object from the name provided as a parameter
    Set MyRange = Range(NameOfCodes)
    On Error GoTo 0

    ' Find the number of rows that are used in the named range
    i = 0
    Do
        i = i + 1
        Value = StrConv(Trim(MyRange.Cells(i, 1).Value), vbUpperCase)
        
        ' Stop when you come to a blank cell
        If Value = "" Then Exit Do
        
        ' Set intRowLast to the row of the last non-blank cell
        intRowLast = i
    Loop

    ' Resize the codes array to hold only as many codes as are needed
    ReDim Codes(1 To intRowLast)

    ' Read the codes into the array, trimming off spaces and converting to uppercase
    For i = 1 To intRowLast
        Codes(i) = StrConv(Trim(MyRange.Cells(i, 1).Value), vbUpperCase)
    Next i
    
    Exit Sub

BadRangeName:
    MsgBox ("Read_Codes is trying to read codes from the range named: " & NameOfCodes & " which does not exist")
    Stop

End Sub

' This function takes a text string from the Psion and determines the line type
' It returns an integer code for line type as defined in the LineTypeEnum
Private Function Get_LineType(ByRef PsionInput As String) As LineTypeEnum
    
    ' Strip off any leading or trailing blanks and convert to all upper case
    PsionInput = Trim(StrConv(PsionInput, vbUpperCase))
    
    ' Handle hard-coded values first
    Select Case Trim(PsionInput)
        Case ""
            Get_LineType = Blank
            Exit Function
        Case "DONE"
            Get_LineType = Done
            Exit Function
        Case "END"
            Get_LineType = FollowEnd
            Exit Function
        Case "FX"
            Get_LineType = FollowEnd
            Exit Function
        Case "GX", "DX"
            Get_LineType = GSEnd
            Exit Function
        Case "CX"
            Get_LineType = FoodPatchEnd
            Exit Function
        Case "VX"
            Get_LineType = AlarmEnd
            Exit Function
        Case "TX"
            Get_LineType = TreeEnd
            Exit Function
        Case "MX"
            Get_LineType = FollowNoMovement
            Exit Function
    End Select
    
    Select Case Left(PsionInput, 2)
        Case "F:", "F."
            If IsLine_FollowStart(PsionInput) Then
                Get_LineType = FollowStart
                Exit Function
            End If
        Case "C "
            Get_LineType = Comment
            Exit Function
        Case "DW", "DR"
            Get_LineType = Other
            Exit Function
    End Select
    
    Select Case Left(PsionInput, 3)
        Case "FS:"
            If IsLine_FollowStart(PsionInput) Then
                Get_LineType = FollowStart
                Exit Function
            End If
        Case "FF."
            If IsLine_FollowStart(PsionInput) Then
                Get_LineType = FollowStart
                Exit Function
            End If
    End Select
    
    If Left(PsionInput, 5) = "ABORT" Then
        Get_LineType = Abort
        Exit Function
    End If
    
    Select Case Left(PsionInput, 1)
        Case "T" ' Any of the Tree codes
            Select Case Mid(PsionInput, 2, 1)
                Case "I"
                    If IsLine_TreeID(PsionInput) Then
                        Get_LineType = TreeID
                        Exit Function
                    End If
                Case "W"
                    If IsLine_Waypoint(PsionInput) Then
                        Get_LineType = TreeWaypoint
                        Exit Function
                    End If
                Case "S"
                    If IsLine_TreeSpecies(PsionInput) Then
                        Get_LineType = TreeSpecies
                        Exit Function
                    End If
                Case "N"
                    If IsLine_TwoWithNumber(PsionInput) Then
                        Get_LineType = TreeNum
                        Exit Function
                    End If
                Case "C"
                    If IsLine_TreeCBH(PsionInput) Then
                        Get_LineType = TreeCBH
                        Exit Function
                    End If
                Case "P"
                    If IsLine_TreePheno(PsionInput) Then
                        Get_LineType = TreePhenology
                        Exit Function
                    End If
                Case "D"
                    If IsLine_TwoWithNumber(PsionInput) Then
                        Get_LineType = TreeDisks
                        Exit Function
                    End If
                Case "B"
                    If IsLine_TwoWithNumber(PsionInput) Then
                        Get_LineType = TreeBromeliads
                        Exit Function
                    End If
            End Select
        
        Case "G", "D"
            If Mid(PsionInput, 2, 2) = "PS" Then
                If IsLine_GPSColor(PsionInput) Then
                    Get_LineType = GPSColor
                    Exit Function
                End If
            End If
            Select Case Mid(PsionInput, 2, 1)
                Case "W"
                    If IsLine_Waypoint(PsionInput) Then
                        Get_LineType = GSWaypoint
                        Exit Function
                    End If
                Case "C"
                    If IsLine_GSClimate(PsionInput) Then
                        Get_LineType = GSClimate
                        Exit Function
                    End If
                Case "L"
                    If IsLine_GSLevel(PsionInput) Then
                        Get_LineType = GSLevel
                        Exit Function
                    End If
                Case "A"
                    If IsLine_GSActivity(PsionInput) Then
                        Get_LineType = GSActivity
                        Exit Function
                    End If
                Case "S"
                    If IsLine_GSStage(PsionInput) Then
                        Get_LineType = GSStage
                        Exit Function
                    End If
                Case "H"
                    If IsLine_GSHeight(PsionInput) Then
                        Get_LineType = GSHeight
                        Exit Function
                    End If
                Case "V"
                    If IsLine_GSVertebrate(PsionInput) Then
                        Get_LineType = GSVertebrate
                        Exit Function
                    End If
            End Select
        
        Case "H"
            Select Case Mid(PsionInput, 2, 1)
                Case "O"
                    If IsLine_HeaderObserver(PsionInput) Then
                        Get_LineType = HeaderObserver
                        Exit Function
                    End If
                Case "G"
                    If IsLine_HeaderGroup(PsionInput) Then
                        Get_LineType = HeaderGroup
                        Exit Function
                    End If
            End Select
            
        Case "W"
            Select Case Mid(PsionInput, 2, 1)
                Case "K"
                    If IsLine_SleepPoint(PsionInput) Then
                        Get_LineType = RangingWake
                        Exit Function
                    End If
                Case "S"
                    If IsLine_SleepPoint(PsionInput) Then
                        Get_LineType = RangingSleep
                        Exit Function
                    End If
                Case "T"
                    If IsLine_RangingWater(PsionInput) Then
                        Get_LineType = RangingWater
                        Exit Function
                    End If
                Case "F"
                    If IsLine_Waypoint(PsionInput) Then
                        Get_LineType = RangingFind
                        Exit Function
                    End If
                Case "L"
                    If IsLine_Waypoint(PsionInput) Then
                        Get_LineType = RangingLeave
                        Exit Function
                    End If
                Case "V"
                    If IsLine_Waypoint(PsionInput) Then
                        Get_LineType = RangingVertebrate
                        Exit Function
                    End If
                Case "."
                    If IsLine_Waypoint(PsionInput) Then
                        Get_LineType = GSWaypoint
                        Exit Function
                    End If
            End Select
        
        Case "Y"
            If IsLine_PSActivity(PsionInput) Then
                Get_LineType = PSActivity
                Exit Function
            End If
        Case "P"
            If IsLine_TwoCode(PsionInput, PostureCodes) Then
                Get_LineType = PSPosture
                Exit Function
            End If
        Case "S"
            If IsLine_TwoCode(PsionInput, SubstrateCodes) Then
                Get_LineType = PSSubstrate
                Exit Function
            End If
        Case "R"
            If IsLine_TwoCode(PsionInput, CentralityCodes) Then
                Get_LineType = PSCentrality
                Exit Function
            End If
        Case "N"
            If IsLine_PSNeighbors(PsionInput) Then
                Get_LineType = PSNeighbors
                Exit Function
            End If
        Case "L"
            If IsLine_PSLevel(PsionInput) Then
                Get_LineType = PSLevel
                Exit Function
            End If
        Case "A"
            If IsLine_TwoCode(PsionInput, ActivityCodes) Then
                Get_LineType = FollowActivity
                Exit Function
            End If
        Case "F"
            If IsLine_TwoCode(PsionInput, SelfDirectedCodes) Then
                Get_LineType = SelfDirected
                Exit Function
            End If
            If IsLine_Waypoint(PsionInput) Then
                Get_LineType = FollowWaypoint
                Exit Function
            End If
        Case "E"
            If IsLine_EatNew(PsionInput) Then
                Get_LineType = EatNew
                Exit Function
            ElseIf IsLine_EatSame(PsionInput) Then
                Get_LineType = EatSame
                Exit Function
            ElseIf IsLine_EatTotal(PsionInput) Then
                Get_LineType = EatTotal
                Exit Function
            End If

        Case "C"
            Select Case Mid(PsionInput, 2, 1)
                Case "F", "I", "B", "P", "S", "R", "A"
                    If IsLine_FoodPatchEnter(PsionInput) Then
                        Get_LineType = FoodPatchEnter
                        Exit Function
                    End If
            End Select
        
        Case "V"
            Select Case PsionInput
                Case "VB", "VU", "VT", "VS"
                    Get_LineType = Alarm
                    Exit Function
            End Select
            Select Case Mid(PsionInput, 2, 1)
                Case "S"
                    If IsLine_AlarmSpecies(PsionInput) Then
                        Get_LineType = AlarmSpecies
                        Exit Function
                    End If
                Case "W"
                    If IsLine_Waypoint(PsionInput) Then
                        Get_LineType = AlarmWaypoint
                        Exit Function
                    End If
                Case "I"
                    If IsLine_AlarmIntensity(PsionInput) Then
                        Get_LineType = AlarmIntensity
                        Exit Function
                    End If
                Case "L"
                    If IsLine_AlarmLevel(PsionInput) Then
                        Get_LineType = AlarmLevel
                        Exit Function
                    End If
                Case "D"
                    If IsLine_AlarmDanger(PsionInput) Then
                        Get_LineType = AlarmDanger
                        Exit Function
                    End If
                Case "M"
                    If IsLine_AlarmMultiple(PsionInput) Then
                        Get_LineType = AlarmMultiple
                        Exit Function
                    End If
                Case "P"
                    If IsLine_AlarmPresent(PsionInput) Then
                        Get_LineType = AlarmPresent
                        Exit Function
                    End If
            End Select
            
        Case "B"
            If IsLine_Behavior(PsionInput) Then
                Get_LineType = Behavior
                Exit Function
            ElseIf IsLine_BehaviorIntergroup(PsionInput) Then
                Get_LineType = BehaviorIntergroup
                Exit Function
            End If
    
        Case "I", "O"
            If IsLine_FollowBlockStatus(PsionInput) Then
                Get_LineType = FollowBlockStatus
                Exit Function
            End If
        
        Case "X"
            If IsLine_FollowStatus(PsionInput) Then
                Get_LineType = FollowStatus
                Exit Function
            End If
        
        Case "Z"
            If IsLine_GPSError(PsionInput) Then
                Get_LineType = GPSError
                Exit Function
            End If

    End Select
    
    Get_LineType = Unknown
    Exit Function

End Function

Private Function IsLine_FollowStart(ByVal PsionInput As String) As Boolean
    IsLine_FollowStart = False
    
    Select Case Len(PsionInput)
        Case 4
            If Not IsCode(MonkeyCodes, Mid(PsionInput, 3)) Then Exit Function
        Case 5
            If Not IsCode(MonkeyCodes, Mid(PsionInput, 4)) Then Exit Function
        Case 10
            If InStr(PsionInput, ".") <> 3 Then Exit Function
            If Not IsCode(FoodCodes, Mid(PsionInput, 4, 4)) Then Exit Function
            If InStr(PsionInput, ":") <> 8 Then Exit Function
            If Not IsCode(MonkeyCodes, Mid(PsionInput, 9)) Then Exit Function
        Case Else
            Exit Function
    End Select
        
    IsLine_FollowStart = True
End Function

Private Function IsLine_Waypoint(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_Waypoint = False
    
    Select Case Mid(PsionInput, 1, 2)
        Case "W."
            If Len(PsionInput) <> 5 Then Exit Function
            If Not StringIsInteger(Mid(PsionInput, 3)) Then Exit Function
        Case Else
            If Len(PsionInput) <> 6 Then Exit Function
            If Mid(PsionInput, 3, 1) <> "." Then Exit Function
            If Not StringIsInteger(Mid(PsionInput, 4)) Then Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_Waypoint = True
End Function

Private Function IsLine_TwoWithNumber(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_TwoWithNumber = False
    
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not StringIsInteger(Mid(PsionInput, 4)) Then Exit Function

    ' PsionInput has passed all tests, so return true
    IsLine_TwoWithNumber = True
End Function

Private Function IsLine_TwoCode(ByVal PsionInput As String, ByRef Codes() As String) As Boolean
    ' Default is false
    IsLine_TwoCode = False
    
    If Len(PsionInput) <> 2 Then Exit Function
    If Not IsCode(Codes, PsionInput) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_TwoCode = True
End Function

Private Function IsLine_TreeID(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_TreeID = False
    
    Dim pos As Integer
    pos = InStr(PsionInput, ".")
    
    If pos <> 3 Then Exit Function
    Select Case Len(Mid(PsionInput, 4))
        Case 1
            If Mid(PsionInput, 4) <> "U" Then Exit Function
        Case 8
            If Not (IsCode(GroupCodes, Mid(PsionInput, 4, 2)) And IsCode(FoodCodes, Mid(PsionInput, 6, 4)) And StringIsInteger(Mid(PsionInput, 10, 2))) And _
            Not (IsCode(FoodCodes, Mid(PsionInput, 4, 4)) And StringIsInteger(Mid(PsionInput, 8, 4))) Then Exit Function
        Case Else
            Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_TreeID = True
End Function

Private Function IsLine_TreeSpecies(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_TreeSpecies = False
    
    If Len(PsionInput) <> 7 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not IsCode(FoodCodes, Mid(PsionInput, 4)) Then Exit Function

    ' PsionInput has passed all tests, so return true
    IsLine_TreeSpecies = True
End Function

Private Function IsLine_TreeCBH(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_TreeCBH = False
    
    Dim substring As String
    Dim Length As Integer
    Dim pos As Integer
    
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    
    substring = Mid(PsionInput, 4)
    
    Do While Len(substring) <> 0
        pos = InStr(substring, ".")
        Select Case pos
            Case 0
                If Not StringIsInteger(substring) Then Exit Function
                substring = ""
            Case Else
                If Not StringIsInteger(Mid(substring, 1, pos - 1)) Then Exit Function
                substring = Mid(substring, pos + 1)
        End Select
    Loop
    
    ' PsionInput has passed all tests, so return true
    IsLine_TreeCBH = True
End Function

Private Function IsLine_TreePheno(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_TreePheno = False
    
    If Len(PsionInput) <> 9 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    
    Select Case Mid(PsionInput, 4, 1)
        Case "0", "1", "2", "3", "4"
        Case Else
            Exit Function
    End Select
    
    Select Case Mid(PsionInput, 5, 1)
        Case "0", "1", "2", "3", "4"
        Case Else
            Exit Function
    End Select
    
    Select Case Mid(PsionInput, 6, 1)
        Case "0", "1", "2", "3", "4"
        Case Else
            Exit Function
    End Select
    
    Select Case Mid(PsionInput, 7, 1)
        Case "0", "1", "2", "3", "4"
        Case Else
            Exit Function
    End Select
    
    Select Case Mid(PsionInput, 8, 1)
        Case "0", "1", "2", "3", "4"
        Case Else
            Exit Function
    End Select
    
    Select Case Mid(PsionInput, 9, 1)
        Case "0", "1", "2", "3", "4"
        Case Else
            Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_TreePheno = True
End Function

Private Function IsLine_GSClimate(ByVal PsionInput As String) As Boolean
    IsLine_GSClimate = False
    If Len(PsionInput) <> 4 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not IsCode(ClimateCodes, Mid(PsionInput, 4)) Then Exit Function
    IsLine_GSClimate = True
End Function

Private Function IsLine_GSLevel(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_GSLevel = False

    Dim Pos1 As Integer, Pos2 As Integer
    Pos1 = InStr(PsionInput, ".")
    Pos2 = InStr(Mid(PsionInput, Pos1 + 1), ".")
    
    If Pos1 <> 3 Then Exit Function
    
    Select Case Len(Mid(PsionInput, 4))
        Case 1
            If Not IsCode(LevelCodes, Mid(PsionInput, 4)) Then Exit Function
        Case 3, 4
            If Not IsCode(LevelCodes, Mid(PsionInput, 4, 1)) Then Exit Function
            If Pos2 <> 2 Then Exit Function
            If CInt(Mid(PsionInput, Pos1 + Pos2 + 1)) Mod 5 <> 0 Then Exit Function
        Case Else
            Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_GSLevel = True
End Function

Private Function IsLine_GSActivity(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_GSActivity = False
    
    Dim Pos1 As Integer, Pos2 As Integer
    Pos1 = InStr(PsionInput, ".")
    Pos2 = InStr(Mid(PsionInput, Pos1 + 1), ".")
    
    If Pos1 <> 3 Then Exit Function
    
    Select Case Len(Mid(PsionInput, 4))
        Case 2
            If Not IsCode(PointSampleCodes, Mid(PsionInput, 4)) Then Exit Function
        Case 7
            If Not IsCode(PointSampleCodes, Mid(PsionInput, 4, 2)) Then Exit Function
            If Pos2 <> 3 Then Exit Function
            If Not IsCode(FoodCodes, Mid(PsionInput, 7)) Then Exit Function
        Case Else
            Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_GSActivity = True
End Function

Private Function IsLine_GSStage(ByVal PsionInput As String) As Boolean
    IsLine_GSStage = False
    If Len(PsionInput) <> 4 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    Select Case Mid(PsionInput, 4)
        Case "0", "1", "2", "3", "4", "5"
        Case Else
            Exit Function
    End Select
    IsLine_GSStage = True
End Function

Private Function IsLine_GSHeight(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_GSHeight = False
    
    If Len(PsionInput) > 5 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not StringIsInteger(Mid(PsionInput, 4)) Then Exit Function
    If CInt(Mid(PsionInput, 4)) Mod 5 <> 0 Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_GSHeight = True
End Function

Private Function IsLine_GSVertebrate(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_GSVertebrate = False
    
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not IsCode(VertebrateCodes, Mid(PsionInput, 4)) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_GSVertebrate = True
End Function

Private Function IsLine_HeaderObserver(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_HeaderObserver = False
    
    Dim Pos1 As Integer
    Pos1 = InStr(PsionInput, ".")
    
    If Pos1 <> 3 Then Exit Function
    If Not IsCode(ObserverCodes, Mid(PsionInput, Pos1 + 1)) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_HeaderObserver = True
End Function

Private Function IsLine_HeaderGroup(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_HeaderGroup = False
    
    Dim Pos1 As Integer
    Pos1 = InStr(PsionInput, ".")
    
    If Pos1 <> 3 Then Exit Function
    If Not IsCode(GroupCodes, Mid(PsionInput, Pos1 + 1)) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_HeaderGroup = True
End Function

Private Function IsLine_SleepPoint(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_SleepPoint = False
    
    If Len(PsionInput) <> 12 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not IsCode(GroupCodes, Mid(PsionInput, 4, 2)) Then Exit Function
    If Mid(PsionInput, 6, 5) <> "SLEEP" Then Exit Function
    If Not StringIsInteger(Mid(PsionInput, 11)) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_SleepPoint = True
End Function

Private Function IsLine_RangingWater(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_RangingWater = False
    
    If Len(PsionInput) <> 12 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not IsCode(GroupCodes, Mid(PsionInput, 4, 2)) Then Exit Function
    If Mid(PsionInput, 6, 5) <> "WATER" Then Exit Function
    If Not StringIsInteger(Mid(PsionInput, 11)) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_RangingWater = True
End Function

Private Function IsLine_PSActivity(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_PSActivity = False
    
    Dim pos As Integer
    Dim Pos2 As Integer
    
    pos = InStr(PsionInput, ".")
    
    If pos <> 0 And pos < 4 Then Exit Function
    
    If Not IsCode(PointSampleCodes, Mid(PsionInput, 2, 2)) Then Exit Function
    
    Select Case Len(PsionInput)
        Case 2
            If StrComp(PsionInput, "Y0") <> 0 Then Exit Function
        Case 3
        Case 4, 8
            If pos <> 4 Then Exit Function
            Select Case Mid(PsionInput, 2, 1)
                Case "F", "E", "X"
                Case Else
                    Exit Function
            End Select
            Select Case Get_FoodObjectIndex(Mid(PsionInput, 3, 1))
                Case FoodObjectEnum.Flower, FoodObjectEnum.Fruit, FoodObjectEnum.Leaf, FoodObjectEnum.Pith, FoodObjectEnum.Seed, FoodObjectEnum.Thorn
                Case Else
                    Exit Function
            End Select
            If Len(PsionInput) = 8 And Not IsCode(FoodCodes, Mid(PsionInput, pos + 1)) Then Exit Function
        Case Else
            Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_PSActivity = True
End Function

Private Function IsLine_PSLevel(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_PSLevel = False

    Dim pos As Integer
    pos = InStr(PsionInput, ".")
    
    Select Case Len(PsionInput)
        Case 2
            If Not IsCode(LevelCodes, Mid(PsionInput, 2)) Then Exit Function
        Case 4, 5
            If Not IsCode(LevelCodes, Mid(PsionInput, 2, 1)) Then Exit Function
            If pos <> 3 Then Exit Function
            If Not StringIsInteger(Mid(PsionInput, 4)) Then Exit Function
            If CInt(Mid(PsionInput, 4)) Mod 5 <> 0 Then Exit Function
        Case Else
            Exit Function
    End Select

    ' PsionInput has passed all tests, so return true
    IsLine_PSLevel = True
End Function

Private Function IsLine_PSNeighbors(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_PSNeighbors = False
    
    Dim ProxString As String
    Dim pos As Integer
    
    If Mid(PsionInput, 2, 2) <> "N." And Mid(PsionInput, 2, 2) <> "D." Then Exit Function
    
    ProxString = Mid(PsionInput, 4)
    If Len(ProxString) < 5 Then Exit Function
    
    pos = InStr(ProxString, "/")
    If pos = 0 Then Exit Function
    If Not StringIsInteger(Left(ProxString, pos - 1)) Then Exit Function
    ProxString = Mid(ProxString, pos + 1)
    pos = InStr(ProxString, "/")
    If pos = 0 Then Exit Function
    If Not StringIsInteger(Left(ProxString, pos - 1)) Then Exit Function
    ProxString = Mid(ProxString, pos + 1)
    If Not StringIsInteger(ProxString) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_PSNeighbors = True
End Function

Private Function IsLine_EatNew(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_EatNew = False

    If Not IsCode(EatCodes, Left(PsionInput, 2)) Then Exit Function
    
    Select Case Len(PsionInput)
        Case 2
        Case 3, 7
            If Mid(PsionInput, 3, 1) <> "." Then Exit Function
            Select Case Get_FoodObjectIndex(Mid(PsionInput, 2, 1))
                Case FoodObjectEnum.Ant, FoodObjectEnum.Bromeliad, FoodObjectEnum.Caterpillar, _
                FoodObjectEnum.Egg, FoodObjectEnum.FoodOther, FoodObjectEnum.Insect, FoodObjectEnum.Nest, _
                FoodObjectEnum.Vertebrate, FoodObjectEnum.Water
                    Exit Function
            End Select
            If Len(PsionInput) = 7 Then
                If Not IsCode(FoodCodes, Right(PsionInput, 4)) Then Exit Function
            End If
        Case Else
            Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_EatNew = True
End Function

Private Function IsLine_EatSame(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_EatSame = False

    If Not IsCode(EatCodes, Left(PsionInput, 2)) Then Exit Function
    
    Select Case Len(PsionInput)
        Case 3
            If Right(PsionInput, 1) <> "/" Then Exit Function
        Case Else
            Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_EatSame = True
End Function

Private Function IsLine_EatTotal(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_EatTotal = False

    If Not IsCode(EatCodes, Left(PsionInput, 2)) Then Exit Function
    
    If Len(PsionInput) > 2 Then
        If Mid(PsionInput, 3, 1) <> "*" Then Exit Function
        If Not StringIsInteger(Mid(PsionInput, 4)) Then Exit Function
    End If
    
    ' PsionInput has passed all tests, so return true
    IsLine_EatTotal = True
End Function

Private Function IsLine_FoodPatchEnter(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_FoodPatchEnter = False

    Dim pos As Integer
    pos = InStr(PsionInput, ".")
    
    Select Case Len(PsionInput)
        Case 2
        Case 3
            If Mid(PsionInput, 3, 1) <> "." Then Exit Function
        Case 7
            If Mid(PsionInput, 3, 1) <> "." Then Exit Function
            If Not IsCode(FoodCodes, Right(PsionInput, 4)) Then Exit Function
            Select Case Get_FoodObjectIndex(Mid(PsionInput, 2, 1))
                Case FoodObjectEnum.Ant, FoodObjectEnum.Caterpillar, _
                FoodObjectEnum.Egg, FoodObjectEnum.FoodOther, FoodObjectEnum.Insect, FoodObjectEnum.Nest, _
                FoodObjectEnum.Vertebrate, FoodObjectEnum.Water
                    Exit Function
            End Select
        Case Else
            Exit Function
    End Select

    ' PsionInput has passed all tests, so return true
    IsLine_FoodPatchEnter = True
End Function

Private Function IsLine_AlarmSpecies(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_AlarmSpecies = False
    
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not IsCode(VertebrateCodes, Mid(PsionInput, 4)) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_AlarmSpecies = True
End Function

Private Function IsLine_AlarmIntensity(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_AlarmIntensity = False
    
    If Len(PsionInput) <> 6 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Mid(PsionInput, 4, 1) <> "O" And Mid(PsionInput, 4, 1) <> "M" And Mid(PsionInput, 4, 1) <> "U" Then Exit Function
    If Mid(PsionInput, 5, 1) <> "O" And Mid(PsionInput, 5, 1) <> "M" And Mid(PsionInput, 5, 1) <> "U" Then Exit Function
    If Mid(PsionInput, 6, 1) <> "J" And Mid(PsionInput, 6, 1) <> "A" And Mid(PsionInput, 6, 1) <> "U" Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_AlarmIntensity = True
End Function

Private Function IsLine_AlarmLevel(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_AlarmLevel = False

    Dim Pos1 As Integer, Pos2 As Integer
    Pos1 = InStr(PsionInput, ".")
    Pos2 = InStr(Mid(PsionInput, Pos1 + 1), ".")

    If Pos1 <> 3 Then Exit Function
    If Not IsCode(LevelCodes, Mid(PsionInput, Pos1 + 1, 1)) Then Exit Function

    If Pos2 <> 0 Then
        If Not StringIsInteger(Mid(PsionInput, Pos1 + Pos2 + 1)) Then Exit Function
        If CInt(Mid(PsionInput, Pos1 + Pos2 + 1)) Mod 5 <> 0 Then Exit Function
    End If

    ' PsionInput has passed all tests, so return true
    IsLine_AlarmLevel = True
End Function

Private Function IsLine_AlarmDanger(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_AlarmDanger = False
    
    If Len(PsionInput) <> 5 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Mid(PsionInput, 4, 1) <> "C" And Mid(PsionInput, 4, 1) <> "N" And Mid(PsionInput, 4, 1) <> "U" Then Exit Function
    If Mid(PsionInput, 5, 1) <> "D" And Mid(PsionInput, 5, 1) <> "N" And Mid(PsionInput, 5, 1) <> "U" Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_AlarmDanger = True
End Function

Private Function IsLine_AlarmMultiple(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_AlarmMultiple = False
    
    If Len(PsionInput) <> 4 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Mid(PsionInput, 4, 1) <> "Y" And Mid(PsionInput, 4, 1) <> "N" And Mid(PsionInput, 4, 1) <> "U" Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_AlarmMultiple = True
End Function

Private Function IsLine_AlarmPresent(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_AlarmPresent = False
    
    If Len(PsionInput) <> 4 Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Mid(PsionInput, 4, 1) <> "Y" And Mid(PsionInput, 4, 1) <> "N" And Mid(PsionInput, 4, 1) <> "U" Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_AlarmPresent = True
End Function

Private Function IsLine_Behavior(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_Behavior = False
    
    If Len(PsionInput) <> 6 Then Exit Function
    If Mid(PsionInput, 2, 1) <> "D" And Mid(PsionInput, 2, 1) <> "S" Then Exit Function
    If Not IsCode(MonkeyCodes, Mid(PsionInput, 3, 2)) Then Exit Function
    If Not IsCode(MonkeyCodes, Mid(PsionInput, 5, 2)) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_Behavior = True
End Function

Private Function IsLine_BehaviorIntergroup(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_BehaviorIntergroup = False
    
    If Len(PsionInput) <> 10 Then Exit Function
    If Mid(PsionInput, 2, 1) <> "I" Then Exit Function
    If Not IsCode(GroupCodes, Mid(PsionInput, 3, 2)) Then Exit Function
    If Mid(PsionInput, 5, 1) <> "." Then Exit Function
    If Mid(PsionInput, 6, 1) <> "W" And Mid(PsionInput, 6, 1) <> "L" And Mid(PsionInput, 6, 1) <> "U" And Mid(PsionInput, 6, 1) <> "D" Then Exit Function
    If Mid(PsionInput, 7, 1) <> "." Then Exit Function
    If Not StringIsInteger(Mid(PsionInput, 8)) Then Exit Function
    
    ' PsionInput has passed all tests, so return true
    IsLine_BehaviorIntergroup = True
End Function

Private Function IsLine_FollowStatus(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_FollowStatus = False

    If Len(PsionInput) <> 2 Then Exit Function
    If Not IsCode(FollowStatusCodes, Mid(PsionInput, 1, 1)) Then Exit Function
    If Not IsCode(FollowPartCodes, Mid(PsionInput, 2, 1)) Then Exit Function
    
    Select Case Mid(PsionInput, 1, 1)
        Case "X"
            If StrComp(Mid(PsionInput, 2, 1), "I") = 0 Or StrComp(Mid(PsionInput, 2, 1), "O") = 0 Then Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_FollowStatus = True
End Function

Private Function IsLine_FollowBlockStatus(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_FollowBlockStatus = False

    If Len(PsionInput) <> 2 Then Exit Function
    If Not IsCode(FollowStatusCodes, Mid(PsionInput, 1, 1)) Then Exit Function
    If Not IsCode(FollowPartCodes, Mid(PsionInput, 2, 1)) Then Exit Function
    
    Select Case Mid(PsionInput, 1, 1)
        Case "O"
            If StrComp(Mid(PsionInput, 2, 1), "I") = 0 Or StrComp(Mid(PsionInput, 2, 1), "X") = 0 Then Exit Function
        Case "I"
            If StrComp(Mid(PsionInput, 2, 1), "O") = 0 Or StrComp(Mid(PsionInput, 2, 1), "X") = 0 Then Exit Function
    End Select
    
    ' PsionInput has passed all tests, so return true
    IsLine_FollowBlockStatus = True
End Function

Private Function IsLine_GPSError(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_GPSError = False

    If Mid(PsionInput, 2, 1) <> "1" And Mid(PsionInput, 2, 1) <> "2" Then Exit Function
    If Mid(PsionInput, 3, 1) <> "." Then Exit Function
    If Not StringIsInteger(Mid(PsionInput, 4)) Then Exit Function

    ' PsionInput has passed all tests, so return true
    IsLine_GPSError = True
End Function

Private Function IsLine_GPSColor(ByVal PsionInput As String) As Boolean
    ' Default is false
    IsLine_GPSColor = False

    If Left(PsionInput, 4) <> "GPS." Then Exit Function
    If Not IsCode(GPSUnits, Mid(PsionInput, 5)) Then Exit Function

    ' PsionInput has passed all tests, so return true
    IsLine_GPSColor = True
End Function

' This function returns the text that corresponds to the given EnumIndex for the given enumeration
' It reads from the range in the "Codes" worksheet
Private Function Get_EnumText(ByVal NameOfEnum As String, ByVal EnumIndex As Integer) As String
    Dim MyRange As Range
    Set MyRange = Range(NameOfEnum)
    
    Get_EnumText = Trim(MyRange.Cells(EnumIndex + 2))

End Function

Private Function Get_FoodObjectIndex(ByVal ObjectCode As String) As FoodObjectEnum
    Select Case ObjectCode
        Case "A"
            Get_FoodObjectIndex = Ant
        Case "B"
            Get_FoodObjectIndex = Bromeliad
        Case "C"
            Get_FoodObjectIndex = Caterpillar
        Case "E"
            Get_FoodObjectIndex = Egg
        Case "F"
            Get_FoodObjectIndex = Fruit
        Case "I"
            Get_FoodObjectIndex = Insect
        Case "L"
            Get_FoodObjectIndex = Leaf
        Case "N"
            Get_FoodObjectIndex = Nest
        Case "O"
            Get_FoodObjectIndex = FoodOther
        Case "P"
            Get_FoodObjectIndex = Pith
        Case "R"
            Get_FoodObjectIndex = Flower
        Case "S"
            Get_FoodObjectIndex = Seed
        Case "T"
            Get_FoodObjectIndex = Thorn
        Case "V"
            Get_FoodObjectIndex = Vertebrate
        Case "W"
            Get_FoodObjectIndex = Water
    End Select
End Function

Private Function Get_WaypointName(ByVal WaypointType As String, ByRef OS As ObservationOutputType) As String
    
    Dim name As String
    
    name = Right(DatePart("yyyy", OS.StartObservation), 2)
    If DatePart("m", OS.StartObservation) < 10 Then
        name = name & "0" & DatePart("m", OS.StartObservation)
    Else
        name = name & DatePart("m", OS.StartObservation)
    End If
    If DatePart("d", OS.StartObservation) < 10 Then
        name = name & "0" & DatePart("d", OS.StartObservation)
    Else
        name = name & DatePart("d", OS.StartObservation)
    End If
    
    Select Case Len(CurrentGPSColor)
        Case 0
            name = WaypointType & "_" & name & "_" & OS.FocalGroup & "_FC_"
        Case 1
            name = WaypointType & "_" & name & "_" & OS.FocalGroup & "_" & CurrentGPSColor & "_FC_"
        Case Else
            MsgBox ("Error with GPS ID code")
    End Select
    Get_WaypointName = name
    
End Function

Private Function Has_SpeciesCode(ByVal TestCode As String) As Boolean
    
     Dim pos As Integer
     
     Has_SpeciesCode = False
     pos = InStr(TestCode, ".")
    
    If pos = 0 Then Exit Function
    If Not IsCode(FoodCodes, Mid(TestCode, pos + 1)) Then Exit Function
    
    Has_SpeciesCode = True

End Function

' This function tests to see if a given strCode is a valid code in the given Codes array
Private Function IsCode(ByRef Codes() As String, ByVal strCode As String) As Boolean
    Dim i As Integer, IMax As Integer
    Dim Code As String
    
    Code = Trim(StrConv(strCode, vbUpperCase))
    IsCode = False
    IMax = UBound(Codes)
    
    For i = LBound(Codes) + 1 To IMax
        ' Check each element of Codes for a match
        If Code = Codes(i) Then
            ' Match found: set IsCode to true and exit function
            IsCode = True
            Exit Function
        End If
    Next i
    
End Function

' Updates the fields for the current Observation
' Called for header codes and ranging events
Private Sub Update_Observation(ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long, ByRef WasLost As Boolean)
            
    Dim i As Long
    
    Select Case InputData(RowIn).LineType
        Case LineTypeEnum.HeaderObserver
            
            If InputData(RowIn).LineType = LineTypeEnum.HeaderObserver Then
                CurrentObservation.Observer = Mid(InputData(RowIn).Data, 4)
            End If
            
            CurrentObservation.ID = CurrentObservation.ID + 1
            CurrentObservation.StartObservation = InputData(RowIn).Datim
            CurrentObservation.FindType = "Find"
            
            ' Find end time of current observation session
            i = 1
            Do Until (RowIn + i) = UBound(InputData)
            ' Loop until end of day
                If InputData(RowIn + i).LineType = LineTypeEnum.Done Or InputData(RowIn + i).LineType = LineTypeEnum.RangingLeave Then Exit Do
                i = i + 1
            Loop
        
            ' Set end time of observation session
            CurrentObservation.EndObservation = InputData(RowIn + i).Datim
            
            ' Make sure end time is after start time
            If CurrentObservation.EndObservation < CurrentObservation.StartObservation Then
                MsgBox "Fatal Error: There is an error beginning on line # " & RowIn & vbCrLf & "Observation End < Observation Begin" & vbCrLf & "Fix the error & parse again"
                Stop
            End If
            
            ' Set the duration
            CurrentObservation.DurationOfObservation = DateDiff("s", CurrentObservation.StartObservation, CurrentObservation.EndObservation)
            
        Case LineTypeEnum.HeaderGroup
            CurrentObservation.FocalGroup = Mid(InputData(RowIn).Data, 4)
        
        Case LineTypeEnum.RangingFind
            CurrentObservation.FindType = "Find"
            CurrentObservation.FindPointID = Get_WaypointName("R", CurrentObservation) & Mid(InputData(RowIn).Data, 4)
            If WasLost Then
                CurrentObservation.ID = CurrentObservation.ID + 1
                CurrentObservation.StartObservation = InputData(RowIn).Datim
                CurrentObservation.FindType = "Find"
                
                ' Find end time of current observation session
                i = 1
                Do Until (RowIn + i) = UBound(InputData)
                ' Loop until end of day
                    If InputData(RowIn + i).LineType = LineTypeEnum.Done Or InputData(RowIn + i).LineType = LineTypeEnum.RangingLeave Then Exit Do
                    i = i + 1
                Loop
            
                ' Set end time of observation session
                CurrentObservation.EndObservation = InputData(RowIn + i).Datim
                
                ' Make sure end time is after start time
                If CurrentObservation.EndObservation < CurrentObservation.StartObservation Then
                    MsgBox "Fatal Error: There is an error beginning on line # " & RowIn & vbCrLf & "Observation End < Observation Begin" & vbCrLf & "Fix the error & parse again"
                    Stop
                End If
                
                ' Set the duration
                CurrentObservation.DurationOfObservation = DateDiff("s", CurrentObservation.StartObservation, CurrentObservation.EndObservation)
            End If
            
            WasLost = False
        Case LineTypeEnum.RangingLeave
            CurrentObservation.LeaveType = "Leave"
            CurrentObservation.LeavePointID = Get_WaypointName("R", CurrentObservation) & Mid(InputData(RowIn).Data, 4)
            WasLost = True
        Case LineTypeEnum.RangingSleep
            CurrentObservation.LeaveType = "Sleep"
            CurrentObservation.LeavePointID = Mid(InputData(RowIn).Data, 4)
            WasLost = True
        Case LineTypeEnum.RangingWake
            CurrentObservation.FindType = "Wake"
            CurrentObservation.FindPointID = Mid(InputData(RowIn).Data, 4)
            WasLost = False
    
    End Select
    
End Sub

' Updates the fields for the current GroupScan
' Called for group scan codes
Private Sub Update_GroupScan(ByRef CurrentGroupScan As GroupScanOutputType, ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    Select Case InputData(RowIn).LineType
        Case LineTypeEnum.GSWaypoint
            CurrentGroupScan.ID = CurrentGroupScan.ID + 1
            CurrentGroupScan.ObservationID = CurrentObservation.ID
            CurrentGroupScan.ScanSeqNum = CurrentGroupScan.ScanSeqNum + 1
            CurrentGroupScan.Datim = InputData(RowIn).Datim
            Select Case Left(InputData(RowIn).Data, 2)
                Case "GW"
                    CurrentGroupScan.WaypointID = Get_WaypointName("R", CurrentObservation) & Mid(InputData(RowIn).Data, 4)
                Case "W."
                    CurrentGroupScan.WaypointID = Get_WaypointName("R", CurrentObservation) & Mid(InputData(RowIn).Data, 3)
                Case Else
                    MsgBox ("Error with waypoint name")
                
            End Select
        Case LineTypeEnum.GSActivity
            CurrentGroupScan.GroupActivity = Mid(InputData(RowIn).Data, 4, 2)
            If InStr(Mid(InputData(RowIn).Data, 4), ".") Then
                CurrentGroupScan.SpeciesCode = Mid(InputData(RowIn).Data, 7)
            Else
                CurrentGroupScan.SpeciesCode = ""
            End If
        Case LineTypeEnum.GSClimate
            CurrentGroupScan.Climate = Mid(InputData(RowIn).Data, 4)
        Case LineTypeEnum.GSHeight
            CurrentGroupScan.CanopyHeight = Mid(InputData(RowIn).Data, 4)
        Case LineTypeEnum.GSLevel
            CurrentGroupScan.ForestLevel = Mid(InputData(RowIn).Data, 4, 1)
            If InStr(Mid(InputData(RowIn).Data, 4), ".") Then
                CurrentGroupScan.GroupHeight = Mid(InputData(RowIn).Data, 6)
            Else
                CurrentGroupScan.GroupHeight = -1
            End If
        Case LineTypeEnum.GSStage
            CurrentGroupScan.Stage = Mid(InputData(RowIn).Data, 4)
    End Select
End Sub

' Updates the fields for the current Vertebrate
Private Sub Update_Vertebrate(ByRef CurrentVertebrate As VertebrateOutputType, ByRef VertebrateData() As VertebrateOutputType, ByRef CurrentGroupScan As GroupScanOutputType, ByRef NextNewVertebrateID As Long, ByRef InputData() As InputLine, ByVal RowIn As Long)
    Dim index As Integer
    
    ' Set ID, foreign keys, and sequence counters
    CurrentVertebrate.ID = NextNewVertebrateID + 1
    CurrentVertebrate.GroupScanID = CurrentGroupScan.ID
    CurrentVertebrate.VertSeqNum = CurrentVertebrate.VertSeqNum + 1
    CurrentVertebrate.Species = Mid(InputData(RowIn).Data, 4)
    
    ' Increment index and redimension the array
    index = UBound(VertebrateData) + 1
    ReDim Preserve VertebrateData(0 To index)
    
    ' Store the current Vertebrate in the array
    VertebrateData(index) = CurrentVertebrate
    
    NextNewVertebrateID = NextNewVertebrateID + 1
        
End Sub

' Updates the fields for the current Follow
' Called for follow begin input lines
Private Sub Update_Follow(ByRef CurrentFollow As FollowOutputType, ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)

    Dim i As Long
    Dim IsFirstPS As Boolean
    Dim CurrentEndTime As Date
    Dim CET_OO As Boolean
    
    IsFirstPS = True
    CET_OO = False
    
    ' Update the ID, sequence numbers, and foreign keys
    CurrentFollow.ID = CurrentFollow.ID + 1
    CurrentFollow.ObservationID = CurrentObservation.ID
    CurrentFollow.SeqNum = CurrentFollow.SeqNum + 1
    
    ' Clear previous comments and other variables
    CurrentFollow.AbortType = ""
    CurrentFollow.Comment = ""
    CurrentFollow.FollowType = ""
    CurrentFollow.WaypointID = ""
    CurrentFollow.SpeciesCode = ""
    
    ' Reset status variables
    CurrentFollow.IsActivityGood = True
    CurrentFollow.IsFollowGood = True
    CurrentFollow.IsForagingGood = True
    CurrentFollow.IsPointGood = True
    CurrentFollow.IsTrackGood = True
    CurrentFollow.IsNoMovement = False
    
    ' Reset error variables
    CurrentFollow.Error1 = -1
    CurrentFollow.Error2 = -1
    CurrentFollow.EatTotal = 1
    
    ' Set start time
    CurrentFollow.StartFollow = InputData(RowIn).Datim
    
    ' Set focal animal and follow type
    Select Case Left(InputData(RowIn).Data, 2)
        Case "F:", "F."
            CurrentFollow.FollowType = "Normal"
            CurrentFollow.FocalAnimal = UCase(Trim(Mid(InputData(RowIn).Data, 3, 2)))
        Case "FS"
            CurrentFollow.FollowType = "Scan"
            CurrentFollow.FocalAnimal = UCase(Trim(Mid(InputData(RowIn).Data, 4, 2)))
        Case "FF"
            CurrentFollow.FollowType = "Feeding"
            CurrentFollow.FocalAnimal = UCase(Trim(Mid(InputData(RowIn).Data, 9, 2)))
            CurrentFollow.SpeciesCode = UCase(Trim(Mid(InputData(RowIn).Data, 4, 4)))
    End Select
    
    CurrentFollow.GPSColor = CurrentGPSColor
    
    ' Find end time of follow and ensure that start time is okay
    i = 1
    Do Until (RowIn + i) = UBound(InputData)
        ' Find the end of the follow: FX, End, or Abort
        Select Case InputData(RowIn + i).LineType
            Case LineTypeEnum.PSActivity
                If IsFirstPS Then
                    CurrentFollow.StartFollow = InputData(RowIn + i).Datim
                    IsFirstPS = False
                End If
                CurrentEndTime = InputData(RowIn + i).Datim
                CET_OO = False
            Case LineTypeEnum.Abort
                CurrentFollow.IsFollowGood = False
                
                ' Take everything after "ABORT" and throw it in the comment
                CurrentFollow.Comment = Trim(Mid(InputData(RowIn + i).Data, 7))
                
                If Not CET_OO Then
                    CurrentEndTime = InputData(RowIn + i).Datim
                End If
                
                ' Test for abort types
                If InStr(InputData(RowIn + i).Data, "LOST") Or Mid(InputData(RowIn + i).Data, 6, 2) = ".L" Then CurrentFollow.AbortType = "LOST"
                If InStr(InputData(RowIn + i).Data, "SWITCH") Or Mid(InputData(RowIn + i).Data, 6, 2) = ".S" Then CurrentFollow.AbortType = "SWITCH"
                If InStr(InputData(RowIn + i).Data, "BAD") Or InStr(InputData(RowIn + i).Data, "WRONG") Or Mid(InputData(RowIn + i).Data, 6, 2) = ".D" Then CurrentFollow.AbortType = "DISCARD"
                Exit Do
            Case LineTypeEnum.FollowEnd
                CurrentFollow.IsFollowGood = True
                CurrentFollow.Comment = Trim(Mid(InputData(RowIn + i).Data, 5))
                Select Case CurrentFollow.FollowType
                    Case "Feeding", "Scan"
                        CurrentEndTime = CurrentFollow.StartFollow
                End Select
                Exit Do
            Case LineTypeEnum.FollowStart
                ' This is an error; new follow begins before old one ends
                MsgBox "Fatal Error: There is no matching END for the follow beginning on input line: " & InputData(RowIn).LineNum
                Exit Sub
            Case LineTypeEnum.FollowWaypoint
                CurrentFollow.WaypointID = Get_WaypointName("S", CurrentObservation) & Mid(InputData(RowIn + i).Data, 4)
            Case LineTypeEnum.EatTotal
                CurrentFollow.EatTotal = CInt(Mid(InputData(RowIn + i).Data, 4))
            Case LineTypeEnum.GPSError
                Select Case Mid(InputData(RowIn + i).Data, 2, 1)
                    Case 1
                        CurrentFollow.Error1 = Mid(InputData(RowIn + i).Data, 4)
                    Case 2
                        CurrentFollow.Error2 = Mid(InputData(RowIn + i).Data, 4)
                End Select
            Case LineTypeEnum.FollowBlockStatus
                Select Case InputData(RowIn + i).Data
                    Case "OO"
                        CurrentEndTime = InputData(RowIn + i).Datim
                        CET_OO = True
                    Case Else
                        CET_OO = False
                End Select
                    
            Case LineTypeEnum.FollowStatus, LineTypeEnum.FollowNoMovement
                Select Case InputData(RowIn + i).Data
                    Case "XT"
                        CurrentFollow.IsTrackGood = False
                    Case "XF"
                        CurrentFollow.IsForagingGood = False
                    Case "XA"
                        CurrentFollow.IsActivityGood = False
                    Case "XP"
                        CurrentFollow.IsPointGood = False
                    Case "MX"
                        CurrentFollow.IsNoMovement = True
                End Select
        End Select
        i = i + 1
    Loop
    
    ' Set end time of follow
    CurrentFollow.EndFollow = CurrentEndTime
    
    ' Make sure end time is after start time
    If CurrentFollow.EndFollow < CurrentFollow.StartFollow Then
        MsgBox "Fatal Error: There is an error in the follow beginning on line # " & RowIn & vbCrLf & "Follow End < Follow Begin" & vbCrLf & "Fix the error & parse again"
        Stop
    End If
    
    ' Set duration of follow
    CurrentFollow.DurationOfFollow = DateDiff("s", CurrentFollow.StartFollow, CurrentFollow.EndFollow)

End Sub

' Updates the fields for the current FollowBlock
' Called for focal begin input lines or any view status change input line
Private Sub Update_FollowBlock(ByRef CurrentFollowBlock As FollowBlockOutputType, ByRef CurrentFollow As FollowOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    Dim i As Long
    Dim CurrentEndTime As Date
    
    ' Test for view status change and set view status
    Select Case InputData(RowIn).LineType
        Case LineTypeEnum.FollowStart
            CurrentFollowBlock.FollowID = CurrentFollow.ID
            CurrentFollowBlock.StartBlock = CurrentFollow.StartFollow
            CurrentFollowBlock.IsInActivity = True
            CurrentFollowBlock.IsInForaging = True
            CurrentFollowBlock.IsInPoint = True
            CurrentFollowBlock.IsInTrack = True
        Case LineTypeEnum.FollowBlockStatus
            CurrentFollowBlock.FollowID = CurrentFollow.ID
            CurrentFollowBlock.StartBlock = InputData(RowIn).Datim
            Select Case InputData(RowIn).Data
                Case "IF"
                    CurrentFollowBlock.IsInForaging = True
                Case "OF"
                    CurrentFollowBlock.IsInForaging = False
                Case "IA"
                    CurrentFollowBlock.IsInActivity = True
                Case "OA"
                    CurrentFollowBlock.IsInActivity = False
                Case "IT"
                    CurrentFollowBlock.IsInTrack = True
                Case "OT"
                    CurrentFollowBlock.IsInTrack = False
                Case "IP"
                    CurrentFollowBlock.IsInPoint = True
                Case "OP"
                    CurrentFollowBlock.IsInPoint = False
                Case "OO"
                    CurrentFollowBlock.IsInActivity = False
                    CurrentFollowBlock.IsInForaging = False
                    CurrentFollowBlock.IsInPoint = False
                    CurrentFollowBlock.IsInTrack = False
                Case "II"
                    CurrentFollowBlock.IsInActivity = True
                    CurrentFollowBlock.IsInForaging = True
                    CurrentFollowBlock.IsInPoint = True
                    CurrentFollowBlock.IsInTrack = True
            End Select
    End Select
    
    ' Find end time of follow block and set initial start time
    i = 1
    Do Until (RowIn + i) = UBound(InputData)
    ' Loop until view state changes or the focal ends
        Select Case InputData(RowIn + i).LineType
            Case LineTypeEnum.Abort, LineTypeEnum.FollowBlockStatus
                CurrentEndTime = InputData(RowIn + i).Datim
                Exit Do
            Case LineTypeEnum.FollowEnd
                CurrentEndTime = CurrentFollow.EndFollow
                Exit Do
        End Select
        i = i + 1
    Loop
    
    ' Set end time of follow block
    CurrentFollowBlock.EndBlock = CurrentEndTime
       
    If CurrentFollowBlock.StartBlock > CurrentFollow.EndFollow Then
        CurrentFollowBlock.StartBlock = CurrentFollow.EndFollow
    End If
    
    If CurrentFollowBlock.EndBlock > CurrentFollow.EndFollow Then
        CurrentFollowBlock.EndBlock = CurrentFollow.EndFollow
    End If
       
    ' Make sure end time is after start time
    If CurrentFollowBlock.EndBlock < CurrentFollowBlock.StartBlock Then
        MsgBox "Fatal Error: There is an error beginning on line # " & RowIn & vbCrLf & "Block End < Block Begin" & vbCrLf & "Fix the error & parse again"
        Stop
    End If
   
   ' Set the duration for the follow block
    CurrentFollowBlock.DurationOfBlock = DateDiff("s", CurrentFollowBlock.StartBlock, CurrentFollowBlock.EndBlock)
    
    If CurrentFollowBlock.DurationOfBlock <> 0 Then CurrentFollowBlock.ID = CurrentFollowBlock.ID + 1

End Sub

' Updates the fields for the current PointSample
' Called for point sample activity lines
Private Sub Update_PointSample(ByRef CurrentPointSample As PointSampleOutputType, ByRef CurrentFollow As FollowOutputType, ByRef CurrentFollowBlock As FollowBlockOutputType, ByRef CurrentSpecies As String, ByRef InputData() As InputLine, ByVal RowIn As Long)
    Dim i As Long
    Dim pos As Integer
    Dim ProxString As String
    
    CurrentPointSample.ID = CurrentPointSample.ID + 1
    CurrentPointSample.FollowID = CurrentFollow.ID
    CurrentPointSample.FollowBlockID = CurrentFollowBlock.ID
    CurrentPointSample.SeqNum = CurrentPointSample.SeqNum + 1
    CurrentPointSample.SpeciesCode = ""
    CurrentPointSample.Datim = InputData(RowIn).Datim
    CurrentPointSample.StateBehav = Mid(InputData(RowIn).Data, 2, 2)
    
    If InStr(InputData(RowIn).Data, ".") Then
        Select Case Len(InputData(RowIn).Data)
            Case 4
                 CurrentPointSample.SpeciesCode = CurrentSpecies
            Case 8
                CurrentSpecies = Mid(InputData(RowIn).Data, 5, 4)
                CurrentPointSample.SpeciesCode = CurrentSpecies
            Case Else
                MsgBox "Fatal Error: There is an error beginning on line # " & RowIn & vbCrLf & "Incorrect point sample code" & vbCrLf & "Fix the error & parse again"
        End Select
    End If
    
    ' Reset other values
    CurrentPointSample.ForestLevel = ""
    CurrentPointSample.Height = -1
    CurrentPointSample.Posture = ""
    CurrentPointSample.Substrate = ""
    CurrentPointSample.Centrality = ""
    CurrentPointSample.IsCarryingDorsal = False
    CurrentPointSample.NumNeighbors0 = -1
    CurrentPointSample.NumNeighbors2 = -1
    CurrentPointSample.NumNeighbors5 = -1
    
    ' Find end time of follow block
    i = 1
    Do Until (RowIn + i) = UBound(InputData)
    ' Loop until all point sample components are recorded
    ' If some components missing, loop until next point sample or end of follow
        Select Case InputData(RowIn + i).LineType
            Case LineTypeEnum.Abort, LineTypeEnum.FollowEnd, LineTypeEnum.PSActivity
                Exit Do
            Case LineTypeEnum.PSLevel
                CurrentPointSample.ForestLevel = Mid(InputData(RowIn + i).Data, 2, 1)
                If InStr(InputData(RowIn + i).Data, ".") Then
                    CurrentPointSample.Height = CInt(Mid(InputData(RowIn + i).Data, 4))
                Else
                    CurrentPointSample.Height = -1
                End If
            Case LineTypeEnum.PSPosture
                CurrentPointSample.Posture = Mid(InputData(RowIn + i).Data, 2, 1)
            Case LineTypeEnum.PSSubstrate
                CurrentPointSample.Substrate = Mid(InputData(RowIn + i).Data, 2, 1)
            Case LineTypeEnum.PSCentrality
                CurrentPointSample.Centrality = Mid(InputData(RowIn + i).Data, 2, 1)
            Case LineTypeEnum.PSNeighbors
                Select Case Left(InputData(RowIn + 1).Data, 2)
                    Case "NN"
                        CurrentPointSample.IsCarryingDorsal = False
                    Case "ND"
                        CurrentPointSample.IsCarryingDorsal = True
                End Select
                
                ProxString = Mid(InputData(RowIn + i).Data, 4)
                pos = InStr(ProxString, "/")
                CurrentPointSample.NumNeighbors0 = CInt(Left(ProxString, pos - 1))
                ProxString = Mid(ProxString, pos + 1)
                pos = InStr(ProxString, "/")
                CurrentPointSample.NumNeighbors2 = CInt(Left(ProxString, pos - 1))
                ProxString = Mid(ProxString, pos + 1)
                CurrentPointSample.NumNeighbors5 = CInt(ProxString)
                
        End Select
        i = i + 1
    Loop
End Sub

' Updates the fields for the current Activity
' Called for activity lines, first point sample of a follow, and II or IA
Private Sub Update_Activity(ByRef CurrentActivity As ActivityOutputType, ByRef CurrentFollow As FollowOutputType, ByRef CurrentFollowBlock As FollowBlockOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    Dim i As Long
    Dim Msg, Style, Response
    
    If InputData(RowIn).Datim > CurrentFollow.EndFollow Or InputData(RowIn).Datim < CurrentFollow.StartFollow Then
        Msg = "Activity code outside of follow at line " & RowIn & ". Quit?"
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Response = MsgBox(Msg, Style)
        If Response = vbYes Then Stop
    End If
    
    CurrentActivity.FollowID = CurrentFollow.ID
    CurrentActivity.FollowBlockID = CurrentFollowBlock.ID
    CurrentActivity.StartState = InputData(RowIn).Datim
    
    Select Case InputData(RowIn).LineType
        Case LineTypeEnum.FollowActivity
            CurrentActivity.Activity = Mid(InputData(RowIn).Data, 2, 1)
        Case LineTypeEnum.PSActivity
            Select Case Mid(InputData(RowIn).Data, 2, 1)
                Case "E", "F", "X"
                    CurrentActivity.Activity = "F"
                Case "A", "D", "U", "R"
                    CurrentActivity.Activity = "R"
                Case "T"
                    CurrentActivity.Activity = "T"
                Case "V"
                    CurrentActivity.Activity = "V"
                Case "S"
                    CurrentActivity.Activity = "S"
                Case "O", "0"
                    CurrentActivity.Activity = "O"
            End Select
        Case LineTypeEnum.FollowBlockStatus
            CurrentActivity.Activity = "O"
    End Select
    
    i = 1
    Do Until (RowIn + i) = UBound(InputData)
    ' Loop until next activity change
        Select Case InputData(RowIn + i).LineType
            Case LineTypeEnum.Abort, LineTypeEnum.FollowEnd, LineTypeEnum.FollowActivity
                Exit Do
            Case LineTypeEnum.FollowBlockStatus
                If InputData(RowIn + i).Data = "OO" Or InputData(RowIn + i).Data = "OA" Then Exit Do
            Case LineTypeEnum.PSActivity
                If CurrentActivity.Activity = "O" Then Exit Do
        End Select
        i = i + 1
    Loop
    
    ' Set end time of activity
    CurrentActivity.EndState = InputData(RowIn + i).Datim
   
    If CurrentActivity.StartState > CurrentFollow.EndFollow Then
        CurrentActivity.StartState = CurrentFollow.EndFollow
    End If

    If CurrentActivity.EndState > CurrentFollow.EndFollow Then
        CurrentActivity.EndState = CurrentFollow.EndFollow
    End If
   
    ' Make sure end time is after start time
    If CurrentActivity.EndState < CurrentActivity.StartState Then
        MsgBox "Fatal Error: There is an error beginning on line # " & RowIn & vbCrLf & "Activity End < Activity Begin" & vbCrLf & "Fix the error & parse again"
        Stop
    End If
   
   ' Set the duration for the activity state
    CurrentActivity.DurationOfState = DateDiff("s", CurrentActivity.StartState, CurrentActivity.EndState)
    
    If CurrentActivity.DurationOfState <> 0 Then CurrentActivity.ID = CurrentActivity.ID + 1
End Sub

' Updates the fields for the current SelfDirected
' Called for self directed lines
Private Sub Update_SelfDirected(ByRef CurrentSelfDirected As SelfDirectedOutputType, ByRef CurrentFollow As FollowOutputType, ByRef CurrentFollowBlock As FollowBlockOutputType, ByRef CurrentActivity As ActivityOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    
    Dim Msg, Style, Response
    
    If InputData(RowIn).Datim > CurrentFollow.EndFollow Or InputData(RowIn).Datim < CurrentFollow.StartFollow Then
        Msg = "Self-directed code outside of follow at line " & RowIn & ". Quit?"
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Response = MsgBox(Msg, Style)
        If Response = vbYes Then Stop
    End If
    
    CurrentSelfDirected.ID = CurrentSelfDirected.ID + 1
    CurrentSelfDirected.FollowID = CurrentFollow.ID
    CurrentSelfDirected.FollowBlockID = CurrentFollowBlock.ID
    CurrentSelfDirected.ActivityID = CurrentActivity.ID
    CurrentSelfDirected.SeqNum = CurrentSelfDirected.SeqNum + 1
    CurrentSelfDirected.Datim = InputData(RowIn).Datim
    CurrentSelfDirected.Behavior = Mid(InputData(RowIn).Data, 2, 1)
End Sub

' Updates the fields for the current FoodPatch
' Called for food patch enter lines
Private Sub Update_FoodPatch(ByRef CurrentFoodPatch As FoodPatchOutputType, ByRef CurrentFollow As FollowOutputType, ByRef CurrentFollowBlock As FollowBlockOutputType, ByRef CurrentSpecies As String, ByRef InputData() As InputLine, ByVal RowIn As Long)
    Dim i As Long
    Dim Msg, Style, Response
    
    If InputData(RowIn).Datim > CurrentFollow.EndFollow Or InputData(RowIn).Datim < CurrentFollow.StartFollow Then
        Msg = "Food patch code outside of follow at line " & RowIn & ". Quit?"
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Response = MsgBox(Msg, Style)
        If Response = vbYes Then Stop
    End If
    
    CurrentFoodPatch.ID = CurrentFoodPatch.ID + 1
    CurrentFoodPatch.FollowID = CurrentFollow.ID
    CurrentFoodPatch.FollowBlockID = CurrentFollowBlock.ID
    
    CurrentFoodPatch.EnterTime = InputData(RowIn).Datim
    CurrentFoodPatch.PatchType = Mid(InputData(RowIn).Data, 2, 1)
    
    Select Case Len(InputData(RowIn).Data)
        Case 2
            CurrentFoodPatch.SpeciesCode = ""
        Case 3
            CurrentFoodPatch.SpeciesCode = CurrentSpecies
        Case 7
            CurrentSpecies = Mid(InputData(RowIn).Data, 4, 4)
            CurrentFoodPatch.SpeciesCode = CurrentSpecies
        Case Else
            MsgBox ("Error: Incorrect length for food patch code on line " & RowIn)
            CurrentFoodPatch.SpeciesCode = ""
    End Select
    
    i = 1
    Do Until (RowIn + i) = UBound(InputData)
    ' Loop until next food patch change
        Select Case InputData(RowIn + i).LineType
            Case LineTypeEnum.Abort, LineTypeEnum.FollowEnd, LineTypeEnum.FoodPatchEnter
                CurrentFoodPatch.IsComplete = False
                Exit Do
            Case LineTypeEnum.FoodPatchEnd
                CurrentFoodPatch.IsComplete = True
                Exit Do
        End Select
        i = i + 1
    Loop
    
    CurrentFoodPatch.ExitTime = InputData(RowIn + i).Datim
    
    If CurrentFoodPatch.EnterTime > CurrentFollow.EndFollow Then
        CurrentFoodPatch.EnterTime = CurrentFollow.EndFollow
    End If

    If CurrentFoodPatch.ExitTime > CurrentFollow.EndFollow Then
        CurrentFoodPatch.ExitTime = CurrentFollow.EndFollow
    End If
    
    CurrentFoodPatch.PatchDuration = DateDiff("s", CurrentFoodPatch.EnterTime, CurrentFoodPatch.ExitTime)
End Sub

' Updates the fields for the current FoodObject
' Called for eat new lines
Private Sub Update_FoodObject(ByRef CurrentFoodObject As FoodObjectOutputType, ByRef CurrentFollow As FollowOutputType, ByRef CurrentFollowBlock As FollowBlockOutputType, ByRef CurrentFoodPatch As FoodPatchOutputType, ByRef CurrentSpecies As String, ByRef IsInFoodPatch As Boolean, ByRef InputData() As InputLine, ByVal RowIn)
    Dim i As Long
    Dim j As Long
    Dim LastBite As Long
    Dim Msg, Style, Response
    
    If InputData(RowIn).Datim > CurrentFollow.EndFollow Or InputData(RowIn).Datim < CurrentFollow.StartFollow Then
        Msg = "Food object code outside of follow at line " & RowIn & ". Quit?"
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Response = MsgBox(Msg, Style)
        If Response = vbYes Then Stop
    End If
    
    CurrentFoodObject.ID = CurrentFoodObject.ID + 1
    CurrentFoodObject.FollowID = CurrentFollow.ID
    CurrentFoodObject.FollowBlockID = CurrentFollowBlock.ID
    If IsInFoodPatch Then
        CurrentFoodObject.FoodPatchID = CurrentFoodPatch.ID
    Else
        CurrentFoodObject.FoodPatchID = -1
    End If
    
    CurrentFoodObject.StartFeeding = InputData(RowIn).Datim
    CurrentFoodObject.FoodItem = Mid(InputData(RowIn).Data, 2, 1)
    
    Select Case Len(InputData(RowIn).Data)
        Case 2
            CurrentFoodObject.SpeciesCode = ""
        Case 3
            CurrentFoodObject.SpeciesCode = CurrentSpecies
        Case 7
            CurrentSpecies = Mid(InputData(RowIn).Data, 4, 4)
            CurrentFoodObject.SpeciesCode = CurrentSpecies
        Case Else
            MsgBox ("Error: Incorrect length for eat code on line " & RowIn)
            CurrentFoodObject.SpeciesCode = ""
    End Select
    
    i = 1
    LastBite = RowIn
    Do Until (RowIn + i) = UBound(InputData)
    ' Loop until next food object change
        Select Case InputData(RowIn + i).LineType
            Case LineTypeEnum.Abort, LineTypeEnum.FollowEnd
                Exit Do
            Case LineTypeEnum.EatNew
                ' Exit if EatNew if for same type of food object
                If StrComp(CurrentFoodObject.FoodItem, Mid(InputData(RowIn + i).Data, 2, 1)) = 0 Then Exit Do
                
                ' Exit unless the next E for the same food item is a EatSame
                j = 1
                Do Until (RowIn + i + j) = UBound(InputData)
                    ' Find next E for same food item
                    Select Case InputData(RowIn + i + j).LineType
                        Case LineTypeEnum.Abort, LineTypeEnum.FollowEnd
                            Exit Do
                        Case LineTypeEnum.EatNew
                            If StrComp(CurrentFoodObject.FoodItem, Mid(InputData(RowIn + i + j).Data, 2, 1)) = 0 Then Exit Do
                        Case LineTypeEnum.EatSame
                            If StrComp(CurrentFoodObject.FoodItem, Mid(InputData(RowIn + i + j).Data, 2, 1)) = 0 Then LastBite = RowIn + i + j
                    End Select
                    j = j + 1
                Loop
                Exit Do
            Case LineTypeEnum.EatSame
                If StrComp(CurrentFoodObject.FoodItem, Mid(InputData(RowIn + i).Data, 2, 1)) = 0 Then LastBite = RowIn + i
        End Select
        i = i + 1
    Loop
    
    CurrentFoodObject.EndFeeding = InputData(LastBite).Datim
    CurrentFoodObject.DurationOfFeeding = DateDiff("s", CurrentFoodObject.StartFeeding, CurrentFoodObject.EndFeeding)
End Sub

' Updates the fields for the current ForagingEvent
' Called for all eat lines (new and same)
Private Sub Update_ForagingEvent(ByRef CurrentForagingEvent As ForagingEventOutputType, ByRef CurrentFollow As FollowOutputType, ByRef CurrentFoodObject As FoodObjectOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    Dim Msg, Style, Response
    
    If InputData(RowIn).Datim > CurrentFollow.EndFollow Or InputData(RowIn).Datim < CurrentFollow.StartFollow Then
        Msg = "Foraging event code outside of follow at line " & RowIn & ". Quit?"
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Response = MsgBox(Msg, Style)
        If Response = vbYes Then Stop
    End If
    
    CurrentForagingEvent.ID = CurrentForagingEvent.ID + 1
    CurrentForagingEvent.FoodObjectID = CurrentFoodObject.ID
    CurrentForagingEvent.SeqNum = CurrentForagingEvent.SeqNum + 1
    CurrentForagingEvent.Datim = InputData(RowIn).Datim
    CurrentForagingEvent.FoodAction = Left(InputData(RowIn).Data, 1)
End Sub

' Updates the fields for the current FruitVisit
' Called for all tree codes
Private Sub Update_FruitVisit(ByRef CurrentFruitVisit As FruitVisitOutputType, ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    
    Select Case InputData(RowIn).LineType
        Case LineTypeEnum.TreeID
            CurrentFruitVisit.ID = CurrentFruitVisit.ID + 1
            CurrentFruitVisit.ObservationID = CurrentObservation.ID
            CurrentFruitVisit.SeqNum = CurrentFruitVisit.SeqNum + 1
            CurrentFruitVisit.Datim = InputData(RowIn).Datim
            CurrentFruitVisit.TreeID = Mid(InputData(RowIn).Data, 4)
            CurrentFruitVisit.SpeciesCode = ""
            CurrentFruitVisit.WaypointID = ""
            CurrentFruitVisit.NumMonkeys = -1
            CurrentFruitVisit.FlowerCover = -1
            CurrentFruitVisit.FlowerMaturity = -1
            CurrentFruitVisit.FruitCover = -1
            CurrentFruitVisit.FruitMaturity = -1
            CurrentFruitVisit.LeafCover = -1
            CurrentFruitVisit.LeafMaturity = -1
            CurrentFruitVisit.NumFruiting = -1
            CurrentFruitVisit.NumPlants = -1
            
            ' Known non-phenology tree
            If IsCode(FoodCodes, Mid(InputData(RowIn).Data, 4, 4)) And IsNumeric(Mid(InputData(RowIn).Data, 8, 4)) Then
                CurrentFruitVisit.SpeciesCode = Mid(InputData(RowIn).Data, 4, 4)
                CurrentFruitVisit.WaypointID = Mid(InputData(RowIn).Data, 4)
            End If
            
            ' Phenology tree
            If IsCode(GroupCodes, Mid(InputData(RowIn).Data, 4, 2)) And IsCode(FoodCodes, Mid(InputData(RowIn).Data, 6, 4)) And IsNumeric(Mid(InputData(RowIn).Data, 10, 2)) Then
                CurrentFruitVisit.SpeciesCode = Mid(InputData(RowIn).Data, 6, 4)
                CurrentFruitVisit.WaypointID = Mid(InputData(RowIn).Data, 4)
            End If
            
            
        Case LineTypeEnum.TreeNum
            CurrentFruitVisit.NumMonkeys = Mid(InputData(RowIn).Data, 4)
        
        Case LineTypeEnum.TreeSpecies
            CurrentFruitVisit.SpeciesCode = Mid(InputData(RowIn).Data, 4)
        
        Case LineTypeEnum.TreeWaypoint
            CurrentFruitVisit.WaypointID = Get_WaypointName("F", CurrentObservation) & Mid(InputData(RowIn).Data, 4)
            
        Case LineTypeEnum.TreeBromeliads
            CurrentFruitVisit.NumPlants = Mid(InputData(RowIn).Data, 4)
        
        Case LineTypeEnum.TreeDisks
            CurrentFruitVisit.NumFruiting = Mid(InputData(RowIn).Data, 4)
        
        Case LineTypeEnum.TreePhenology
            CurrentFruitVisit.LeafCover = CInt(Mid(InputData(RowIn).Data, 4, 1))
            CurrentFruitVisit.LeafMaturity = CInt(Mid(InputData(RowIn).Data, 5, 1))
            CurrentFruitVisit.FruitCover = CInt(Mid(InputData(RowIn).Data, 6, 1))
            CurrentFruitVisit.FruitMaturity = CInt(Mid(InputData(RowIn).Data, 7, 1))
            CurrentFruitVisit.FlowerCover = CInt(Mid(InputData(RowIn).Data, 8, 1))
            CurrentFruitVisit.FlowerMaturity = CInt(Mid(InputData(RowIn).Data, 9, 1))
            
    End Select
    
End Sub

' Updates the fields for the current TreeCBH
' Called for tree CBH lines
Private Sub Update_TreeCBH(ByRef CurrentTreeCBH As TreeCBHOutputType, ByRef CurrentFruitVisit As FruitVisitOutputType, ByVal stemCBH As Integer)
    CurrentTreeCBH.ID = CurrentTreeCBH.ID + 1
    CurrentTreeCBH.StemNum = CurrentTreeCBH.StemNum + 1
    CurrentTreeCBH.FruitVisitID = CurrentFruitVisit.ID
    CurrentTreeCBH.CBH = stemCBH
End Sub

' Updates the fields for the current Alarm
' Called for all alarm codes
Private Sub Update_Alarm(ByRef CurrentAlarm As AlarmOutputType, ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn)
    
    Select Case InputData(RowIn).LineType
        Case LineTypeEnum.Alarm
            CurrentAlarm.ID = CurrentAlarm.ID + 1
            CurrentAlarm.ObservationID = CurrentObservation.ID
            CurrentAlarm.PredatorType = Mid(InputData(RowIn).Data, 2, 1)
            CurrentAlarm.Datim = InputData(RowIn).Datim
            
            ' Reset fields
            CurrentAlarm.AlarmerAge = ""
            CurrentAlarm.Danger = ""
            CurrentAlarm.ForestLevel = ""
            CurrentAlarm.Height = -1
            CurrentAlarm.IsConfirmed = False
            CurrentAlarm.IsMultiple = False
            CurrentAlarm.IsPresent = False
            CurrentAlarm.NumAlarmers = ""
            CurrentAlarm.NumAlarms = ""
            CurrentAlarm.PredatorSpecies = ""
        Case LineTypeEnum.AlarmDanger
            Select Case Mid(InputData(RowIn).Data, 4, 1)
                Case "C"
                    CurrentAlarm.IsConfirmed = True
                Case "N", "U"
                    CurrentAlarm.IsConfirmed = False
            End Select
            CurrentAlarm.Danger = Mid(InputData(RowIn).Data, 5, 1)
        Case LineTypeEnum.AlarmIntensity
            CurrentAlarm.NumAlarmers = Mid(InputData(RowIn).Data, 4, 1)
            CurrentAlarm.NumAlarms = Mid(InputData(RowIn).Data, 5, 1)
            CurrentAlarm.AlarmerAge = Mid(InputData(RowIn).Data, 6, 1)
        Case LineTypeEnum.AlarmLevel
            CurrentAlarm.ForestLevel = Mid(InputData(RowIn).Data, 4, 1)
            If InStr(Mid(InputData(RowIn).Data, 4), ".") Then
                CurrentAlarm.Height = Mid(InputData(RowIn).Data, 6)
            Else
                CurrentAlarm.Height = -1
            End If
        Case LineTypeEnum.AlarmMultiple
            Select Case Mid(InputData(RowIn).Data, 4, 1)
                Case "Y"
                    CurrentAlarm.IsMultiple = True
                Case "N"
                    CurrentAlarm.IsMultiple = False
            End Select
        Case LineTypeEnum.AlarmPresent
            Select Case Mid(InputData(RowIn).Data, 4, 1)
                Case "Y"
                    CurrentAlarm.IsPresent = True
                Case "N"
                    CurrentAlarm.IsPresent = False
            End Select
        Case LineTypeEnum.AlarmSpecies
            CurrentAlarm.PredatorSpecies = Mid(InputData(RowIn).Data, 4)
        Case LineTypeEnum.AlarmWaypoint
            CurrentAlarm.WaypointID = Get_WaypointName("A", CurrentObservation) & Mid(InputData(RowIn).Data, 4)
    End Select
End Sub

' Updates the fields for the current Interaction
' Called for interaction codes
Private Sub Update_Interaction(ByRef CurrentInteraction As InteractionOutputType, ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    CurrentInteraction.ID = CurrentInteraction.ID + 1
    CurrentInteraction.ObservationID = CurrentObservation.ID
    CurrentInteraction.Datim = InputData(RowIn).Datim
    CurrentInteraction.InteractionType = Mid(InputData(RowIn).Data, 2, 1)
    CurrentInteraction.Actor = Mid(InputData(RowIn).Data, 3, 2)
    CurrentInteraction.Recipient = Mid(InputData(RowIn).Data, 5, 2)
End Sub

' Updates the fields for the current Intergroup
' Called for intergroup codes
Private Sub Update_Intergroup(ByRef CurrentIntergroup As IntergroupOutputType, ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    CurrentIntergroup.ID = CurrentIntergroup.ID + 1
    CurrentIntergroup.ObservationID = CurrentObservation.ID
    CurrentIntergroup.Datim = InputData(RowIn).Datim
    CurrentIntergroup.OpponentGroup = Mid(InputData(RowIn).Data, 3, 2)
    CurrentIntergroup.Outcome = Mid(InputData(RowIn).Data, 6, 1)
    CurrentIntergroup.WaypointID = Get_WaypointName("I", CurrentObservation) & Right(InputData(RowIn).Data, 3)
End Sub

' Updates the fields for the current RangingEvent
' Called for water and vertebrate eating codes
Private Sub Update_RangingEvent(ByRef CurrentRangingEvent As RangingEventOutputType, ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn As Long)
    CurrentRangingEvent.ID = CurrentRangingEvent.ID + 1
    CurrentRangingEvent.ObservationID = CurrentObservation.ID
    CurrentRangingEvent.Datim = InputData(RowIn).Datim
    
    Select Case Mid(InputData(RowIn).Data, 2, 1)
        Case "T"
            CurrentRangingEvent.EventType = "Water"
            CurrentRangingEvent.WaypointID = Mid(InputData(RowIn).Data, 4)
        Case "V"
            CurrentRangingEvent.EventType = "Vertebrate"
            CurrentRangingEvent.WaypointID = Get_WaypointName("V", CurrentObservation) & Mid(InputData(RowIn).Data, 4)
    End Select
    
End Sub

' Updates the fields for the current Comment
' Called for comment codes
Private Sub Update_Comment(ByRef CurrentComment As CommentOutputType, ByRef CurrentObservation As ObservationOutputType, ByRef InputData() As InputLine, ByVal RowIn)
    CurrentComment.ID = CurrentComment.ID + 1
    CurrentComment.ObservationID = CurrentObservation.ID
    CurrentComment.Datim = InputData(RowIn).Datim
    CurrentComment.Comment = InputData(RowIn).Data
End Sub

Private Function Parse_TreeCBH(ByVal PsionInput As String) As Integer()
    Dim Output() As Integer
    Dim c As String
    Dim pos As Integer
    Dim numStems As Integer
    Dim i As Integer
   
    c = Mid(PsionInput, 4)
    numStems = 0
    
    Do Until Len(c) = 0
        numStems = numStems + 1
        ReDim Preserve Output(1 To numStems)
        pos = InStr(c, ".")
        Select Case pos
            Case 0
                Output(numStems) = c
                c = ""
            Case Else
                Output(numStems) = Left(c, pos - 1)
                c = Mid(c, pos + 1)
        End Select
    Loop
    
   Parse_TreeCBH = Output
End Function

Private Sub Write_InputError_Header()
    ' Create a new worksheet for input errors and make it the active worksheet
    Call Worksheet_Create("InputError")

    Cells(1, 1).Value = "InputLine"
    Cells(1, 2).Value = "Psion Date/Time"
    Cells(1, 3).Value = "Psion Data"
    Cells(1, 4).Value = "Type of Error"
End Sub

Private Sub Write_InputError(ByVal ErrorNumber As Integer, ByRef InputData As InputLine, ByVal strError As String)
    Cells(ErrorNumber + 1, 1).Value = InputData.LineNum
    Cells(ErrorNumber + 1, 2).Value = InputData.Datim
    Cells(ErrorNumber + 1, 3).Value = InputData.Data
    Cells(ErrorNumber + 1, 4).Value = strError
End Sub

Private Sub Write_ObservationHeader()
    With Worksheets("Observation")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "FocalGroup"
        .Range("C1").Value = "Observer"
        .Range("D1").Value = "StartObservation"
        .Range("E1").Value = "EndObservation"
        .Range("F1").Value = "DurationOfObservation"
        .Range("G1").Value = "FindType"
        .Range("H1").Value = "FindPointID"
        .Range("I1").Value = "LeaveType"
        .Range("J1").Value = "LeavePointID"
        .Range("K1").Value = "IsFullDay"
    End With
End Sub

Private Sub Write_Observation(ByRef Observation As ObservationOutputType)
    Dim RowOut As Long
    RowOut = Observation.ID + 1
    With Worksheets("Observation")
        .Range("A" & RowOut).Value = Observation.ID
        .Range("B" & RowOut).Value = Observation.FocalGroup
        .Range("C" & RowOut).Value = Observation.Observer
        .Range("D" & RowOut).Value = Observation.StartObservation
        .Range("E" & RowOut).Value = Observation.EndObservation
        .Range("F" & RowOut).Value = Observation.DurationOfObservation
        .Range("G" & RowOut).Value = Observation.FindType
        .Range("H" & RowOut).Value = Observation.FindPointID
        .Range("I" & RowOut).Value = Observation.LeaveType
        .Range("J" & RowOut).Value = Observation.LeavePointID
        .Range("K" & RowOut).Value = Observation.IsFullDay
    End With
End Sub

Private Sub Write_GroupScanHeader()
    With Worksheets("GroupScan")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ObservationID"
        .Range("C1").Value = "WaypointID"
        .Range("D1").Value = "ScanSeqNum"
        .Range("E1").Value = "Datim"
        .Range("F1").Value = "GroupActivity"
        .Range("G1").Value = "SpeciesCode"
        .Range("H1").Value = "ForestLevel"
        .Range("I1").Value = "CanopyHeight"
        .Range("J1").Value = "GroupHeight"
        .Range("K1").Value = "Climate"
        .Range("L1").Value = "Stage"
    End With
End Sub

Private Sub Write_GroupScan(ByRef GroupScan As GroupScanOutputType)
    Dim RowOut As Long
    RowOut = GroupScan.ID + 1
    With Worksheets("GroupScan")
        .Range("A" & RowOut).Value = GroupScan.ID
        .Range("B" & RowOut).Value = GroupScan.ObservationID
        .Range("C" & RowOut).Value = GroupScan.WaypointID
        .Range("D" & RowOut).Value = GroupScan.ScanSeqNum
        .Range("E" & RowOut).Value = GroupScan.Datim
        .Range("F" & RowOut).Value = GroupScan.GroupActivity
        .Range("G" & RowOut).Value = GroupScan.SpeciesCode
        .Range("H" & RowOut).Value = GroupScan.ForestLevel
        If GroupScan.CanopyHeight <> -1 Then .Range("I" & RowOut).Value = GroupScan.CanopyHeight
        If GroupScan.GroupHeight <> -1 Then .Range("J" & RowOut).Value = GroupScan.GroupHeight
        .Range("K" & RowOut).Value = GroupScan.Climate
        If GroupScan.Stage <> -1 Then .Range("L" & RowOut).Value = GroupScan.Stage
    End With
End Sub

Private Sub Write_VertebrateHeader()
    With Worksheets("Vertebrate")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "GroupScanID"
        .Range("C1").Value = "VertSeqNum"
        .Range("D1").Value = "Species"
    End With
End Sub

Private Sub Write_Vertebrate(ByRef Vertebrate As VertebrateOutputType)
    Dim RowOut As Long
    RowOut = Vertebrate.ID + 1
    With Worksheets("Vertebrate")
        .Range("A" & RowOut).Value = Vertebrate.ID
        .Range("B" & RowOut).Value = Vertebrate.GroupScanID
        .Range("C" & RowOut).Value = Vertebrate.VertSeqNum
        .Range("D" & RowOut).Value = Vertebrate.Species
    End With
End Sub

Private Sub Write_FollowHeader()
    With Worksheets("Follow")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ObservationID"
        .Range("C1").Value = "SeqNum"
        .Range("D1").Value = "FocalAnimal"
        .Range("E1").Value = "StartFollow"
        .Range("F1").Value = "EndFollow"
        .Range("G1").Value = "DurationOfFollow"
        .Range("H1").Value = "FollowType"
        .Range("I1").Value = "SpeciesCode"
        .Range("J1").Value = "WaypointID"
        .Range("K1").Value = "IsFollowGood"
        .Range("L1").Value = "IsTrackGood"
        .Range("M1").Value = "IsPointGood"
        .Range("N1").Value = "IsActivityGood"
        .Range("O1").Value = "IsForagingGood"
        .Range("P1").Value = "IsNoMovement"
        .Range("Q1").Value = "GPSColor"
        .Range("R1").Value = "Error1"
        .Range("S1").Value = "Error2"
        .Range("T1").Value = "EatTotal"
        .Range("U1").Value = "AbortType"
        .Range("V1").Value = "Comment"
    End With
End Sub

Private Sub Write_Follow(ByRef Follow As FollowOutputType)
    Dim RowOut As Long
    RowOut = Follow.ID + 1
    With Worksheets("Follow")
        .Range("A" & RowOut).Value = Follow.ID
        .Range("B" & RowOut).Value = Follow.ObservationID
        .Range("C" & RowOut).Value = Follow.SeqNum
        .Range("D" & RowOut).Value = Follow.FocalAnimal
        .Range("E" & RowOut).Value = Follow.StartFollow
        .Range("F" & RowOut).Value = Follow.EndFollow
        .Range("G" & RowOut).Value = Follow.DurationOfFollow
        .Range("H" & RowOut).Value = Follow.FollowType
        .Range("I" & RowOut).Value = Follow.SpeciesCode
        .Range("J" & RowOut).Value = Follow.WaypointID
        If Follow.FollowType = "Normal" Then
            .Range("K" & RowOut).Value = Follow.IsFollowGood
            .Range("L" & RowOut).Value = Follow.IsTrackGood
            .Range("M" & RowOut).Value = Follow.IsPointGood
            .Range("N" & RowOut).Value = Follow.IsActivityGood
            .Range("O" & RowOut).Value = Follow.IsForagingGood
            .Range("P" & RowOut).Value = Follow.IsNoMovement
        End If
        .Range("Q" & RowOut).Value = Follow.GPSColor
        If Follow.Error1 <> -1 Then .Range("R" & RowOut).Value = Follow.Error1
        If Follow.Error2 <> -1 Then .Range("S" & RowOut).Value = Follow.Error2
        If Follow.FollowType = "Feeding" Then
            .Range("T" & RowOut).Value = Follow.EatTotal
        End If
        .Range("U" & RowOut).Value = Follow.AbortType
        .Range("V" & RowOut).Value = Follow.Comment
    End With
End Sub

Private Sub Write_FollowBlockHeader()
    With Worksheets("FollowBlock")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "FollowID"
        .Range("C1").Value = "StartBlock"
        .Range("D1").Value = "EndBlock"
        .Range("E1").Value = "DurationOfBlock"
        .Range("F1").Value = "IsInTrack"
        .Range("G1").Value = "IsInForaging"
        .Range("H1").Value = "IsInActivity"
        .Range("I1").Value = "IsInPoint"
    End With
End Sub

Private Sub Write_FollowBlock(ByRef FollowBlock As FollowBlockOutputType)
    Dim RowOut As Long
    RowOut = FollowBlock.ID + 1
    With Worksheets("FollowBlock")
        .Range("A" & RowOut).Value = FollowBlock.ID
        .Range("B" & RowOut).Value = FollowBlock.FollowID
        .Range("C" & RowOut).Value = FollowBlock.StartBlock
        .Range("D" & RowOut).Value = FollowBlock.EndBlock
        .Range("E" & RowOut).Value = FollowBlock.DurationOfBlock
        .Range("F" & RowOut).Value = FollowBlock.IsInTrack
        .Range("G" & RowOut).Value = FollowBlock.IsInForaging
        .Range("H" & RowOut).Value = FollowBlock.IsInActivity
        .Range("I" & RowOut).Value = FollowBlock.IsInPoint
    End With
End Sub

Private Sub Write_PointSampleHeader()
    With Worksheets("PointSample")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "FollowID"
        .Range("C1").Value = "FollowBlockID"
        .Range("D1").Value = "SeqNum"
        .Range("E1").Value = "Datim"
        .Range("F1").Value = "StateBehav"
        .Range("G1").Value = "SpeciesCode"
        .Range("H1").Value = "Posture"
        .Range("I1").Value = "Substrate"
        .Range("J1").Value = "ForestLevel"
        .Range("K1").Value = "Height"
        .Range("L1").Value = "Centrality"
        .Range("M1").Value = "IsCarryingDorsal"
        .Range("N1").Value = "NumNeighbors0"
        .Range("O1").Value = "NumNeighbors2"
        .Range("P1").Value = "NumNeighbors5"
    End With
End Sub

Private Sub Write_PointSample(ByRef PointSample As PointSampleOutputType)
    Dim RowOut As Long
    RowOut = PointSample.ID + 1
    With Worksheets("PointSample")
        .Range("A" & RowOut).Value = PointSample.ID
        .Range("B" & RowOut).Value = PointSample.FollowID
        .Range("C" & RowOut).Value = PointSample.FollowBlockID
        .Range("D" & RowOut).Value = PointSample.SeqNum
        .Range("E" & RowOut).Value = PointSample.Datim
        .Range("F" & RowOut).Value = PointSample.StateBehav
        .Range("G" & RowOut).Value = PointSample.SpeciesCode
        .Range("H" & RowOut).Value = PointSample.Posture
        .Range("I" & RowOut).Value = PointSample.Substrate
        .Range("J" & RowOut).Value = PointSample.ForestLevel
        If PointSample.Height <> -1 Then .Range("K" & RowOut).Value = PointSample.Height
        .Range("L" & RowOut).Value = PointSample.Centrality
        If PointSample.NumNeighbors0 <> -1 Then .Range("M" & RowOut).Value = PointSample.IsCarryingDorsal
        If PointSample.NumNeighbors0 <> -1 Then .Range("N" & RowOut).Value = PointSample.NumNeighbors0
        If PointSample.NumNeighbors2 <> -1 Then .Range("O" & RowOut).Value = PointSample.NumNeighbors2
        If PointSample.NumNeighbors5 <> -1 Then .Range("P" & RowOut).Value = PointSample.NumNeighbors5
    End With
End Sub

Private Sub Write_ActivityHeader()
    With Worksheets("Activity")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "FollowID"
        .Range("C1").Value = "FollowBlockID"
        .Range("D1").Value = "Activity"
        .Range("E1").Value = "StartState"
        .Range("F1").Value = "EndState"
        .Range("G1").Value = "DurationOfState"
    End With
End Sub

Private Sub Write_Activity(ByRef Activity As ActivityOutputType)
    Dim RowOut As Long
    RowOut = Activity.ID + 1
    With Worksheets("Activity")
        .Range("A" & RowOut).Value = Activity.ID
        .Range("B" & RowOut).Value = Activity.FollowID
        .Range("C" & RowOut).Value = Activity.FollowBlockID
        .Range("D" & RowOut).Value = Activity.Activity
        .Range("E" & RowOut).Value = Activity.StartState
        .Range("F" & RowOut).Value = Activity.EndState
        .Range("G" & RowOut).Value = Activity.DurationOfState
    End With
End Sub

Private Sub Write_SelfDirectedHeader()
    With Worksheets("SelfDirected")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "FollowID"
        .Range("C1").Value = "FollowBlockID"
        .Range("D1").Value = "ActivityID"
        .Range("E1").Value = "Datim"
        .Range("F1").Value = "SeqNum"
        .Range("G1").Value = "Behavior"
    End With
End Sub

Private Sub Write_SelfDirected(ByRef SelfDirected As SelfDirectedOutputType)
    Dim RowOut As Long
    RowOut = SelfDirected.ID + 1
    With Worksheets("SelfDirected")
        .Range("A" & RowOut).Value = SelfDirected.ID
        .Range("B" & RowOut).Value = SelfDirected.FollowID
        .Range("C" & RowOut).Value = SelfDirected.FollowBlockID
        .Range("D" & RowOut).Value = SelfDirected.ActivityID
        .Range("E" & RowOut).Value = SelfDirected.Datim
        .Range("F" & RowOut).Value = SelfDirected.SeqNum
        .Range("G" & RowOut).Value = SelfDirected.Behavior
    End With
End Sub

Private Sub Write_FoodPatchHeader()
    With Worksheets("FoodPatch")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "FollowID"
        .Range("C1").Value = "FollowBlockID"
        .Range("D1").Value = "EnterTime"
        .Range("E1").Value = "ExitTime"
        .Range("F1").Value = "PatchType"
        .Range("G1").Value = "SpeciesCode"
        .Range("H1").Value = "PatchDuration"
        .Range("I1").Value = "IsComplete"
    End With
End Sub

Private Sub Write_FoodPatch(ByRef FoodPatch As FoodPatchOutputType)
    Dim RowOut As Long
    RowOut = FoodPatch.ID + 1
    With Worksheets("FoodPatch")
        .Range("A" & RowOut).Value = FoodPatch.ID
        .Range("B" & RowOut).Value = FoodPatch.FollowID
        .Range("C" & RowOut).Value = FoodPatch.FollowBlockID
        .Range("D" & RowOut).Value = FoodPatch.EnterTime
        .Range("E" & RowOut).Value = FoodPatch.ExitTime
        .Range("F" & RowOut).Value = FoodPatch.PatchType
        .Range("G" & RowOut).Value = FoodPatch.SpeciesCode
        .Range("H" & RowOut).Value = FoodPatch.PatchDuration
        .Range("I" & RowOut).Value = FoodPatch.IsComplete
    End With
End Sub

Private Sub Write_FoodObjectHeader()
    With Worksheets("FoodObject")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "FollowID"
        .Range("C1").Value = "FollowBlockID"
        .Range("D1").Value = "FoodPatchID"
        .Range("E1").Value = "FoodItem"
        .Range("F1").Value = "SpeciesCode"
        .Range("G1").Value = "StartFeeding"
        .Range("H1").Value = "EndFeeding"
        .Range("I1").Value = "DurationOfFeeding"
    End With
End Sub

Private Sub Write_FoodObject(ByRef FoodObject As FoodObjectOutputType)
    Dim RowOut As Long
    RowOut = FoodObject.ID + 1
    With Worksheets("FoodObject")
        .Range("A" & RowOut).Value = FoodObject.ID
        .Range("B" & RowOut).Value = FoodObject.FollowID
        .Range("C" & RowOut).Value = FoodObject.FollowBlockID
        If FoodObject.FoodPatchID <> -1 Then .Range("D" & RowOut).Value = FoodObject.FoodPatchID
        .Range("E" & RowOut).Value = FoodObject.FoodItem
        .Range("F" & RowOut).Value = FoodObject.SpeciesCode
        .Range("G" & RowOut).Value = FoodObject.StartFeeding
        .Range("H" & RowOut).Value = FoodObject.EndFeeding
        .Range("I" & RowOut).Value = FoodObject.DurationOfFeeding
    End With
End Sub

Private Sub Write_ForagingEventHeader()
    With Worksheets("ForagingEvent")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "FoodObjectID"
        .Range("C1").Value = "Datim"
        .Range("D1").Value = "SeqNum"
        .Range("E1").Value = "FoodAction"
    End With
End Sub

Private Sub Write_ForagingEvent(ByRef ForagingEvent As ForagingEventOutputType)
    Dim RowOut As Long
    RowOut = ForagingEvent.ID + 1
    With Worksheets("ForagingEvent")
        .Range("A" & RowOut).Value = ForagingEvent.ID
        .Range("B" & RowOut).Value = ForagingEvent.FoodObjectID
        .Range("C" & RowOut).Value = ForagingEvent.Datim
        .Range("D" & RowOut).Value = ForagingEvent.SeqNum
        .Range("E" & RowOut).Value = ForagingEvent.FoodAction
    End With
End Sub

Private Sub Write_FruitVisitHeader()
    With Worksheets("FruitVisit")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ObservationID"
        .Range("C1").Value = "TreeID"
        .Range("D1").Value = "WaypointID"
        .Range("E1").Value = "SeqNum"
        .Range("F1").Value = "Datim"
        .Range("G1").Value = "SpeciesCode"
        .Range("H1").Value = "NumMonkeys"
        .Range("I1").Value = "LeafCover"
        .Range("J1").Value = "LeafMaturity"
        .Range("K1").Value = "FruitCover"
        .Range("L1").Value = "FruitMaturity"
        .Range("M1").Value = "FlowerCover"
        .Range("N1").Value = "FlowerMaturity"
        .Range("O1").Value = "NumPlants"
        .Range("P1").Value = "NumFruiting"
    End With
End Sub

Private Sub Write_FruitVisit(ByRef FruitVisit As FruitVisitOutputType)
    Dim RowOut As Long
    RowOut = FruitVisit.ID + 1
    With Worksheets("FruitVisit")
        .Range("A" & RowOut).Value = FruitVisit.ID
        .Range("B" & RowOut).Value = FruitVisit.ObservationID
        .Range("C" & RowOut).Value = FruitVisit.TreeID
        .Range("D" & RowOut).Value = FruitVisit.WaypointID
        .Range("E" & RowOut).Value = FruitVisit.SeqNum
        .Range("F" & RowOut).Value = FruitVisit.Datim
        .Range("G" & RowOut).Value = FruitVisit.SpeciesCode
        If FruitVisit.NumMonkeys <> -1 Then .Range("H" & RowOut).Value = FruitVisit.NumMonkeys
        If FruitVisit.LeafCover <> -1 Then .Range("I" & RowOut).Value = FruitVisit.LeafCover
        If FruitVisit.LeafMaturity <> -1 Then .Range("J" & RowOut).Value = FruitVisit.LeafMaturity
        If FruitVisit.FruitCover <> -1 Then .Range("K" & RowOut).Value = FruitVisit.FruitCover
        If FruitVisit.FruitMaturity <> -1 Then .Range("L" & RowOut).Value = FruitVisit.FruitMaturity
        If FruitVisit.FlowerCover <> -1 Then .Range("M" & RowOut).Value = FruitVisit.FlowerCover
        If FruitVisit.FlowerMaturity <> -1 Then .Range("N" & RowOut).Value = FruitVisit.FlowerMaturity
        If FruitVisit.NumPlants <> -1 Then .Range("O" & RowOut).Value = FruitVisit.NumPlants
        If FruitVisit.NumFruiting <> -1 Then .Range("P" & RowOut).Value = FruitVisit.NumFruiting
    End With
End Sub

Private Sub Write_TreeCBHHeader()
    With Worksheets("TreeCBH")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "TreeID"
        .Range("C1").Value = "FruitVisitID"
        .Range("D1").Value = "StemNum"
        .Range("E1").Value = "CBH"
    End With
End Sub

Private Sub Write_TreeCBH(ByRef TreeCBH As TreeCBHOutputType)
    Dim RowOut As Long
    RowOut = TreeCBH.ID + 1
    With Worksheets("TreeCBH")
        .Range("A" & RowOut).Value = TreeCBH.ID
        .Range("B" & RowOut).Value = TreeCBH.TreeID
        .Range("C" & RowOut).Value = TreeCBH.FruitVisitID
        .Range("D" & RowOut).Value = TreeCBH.StemNum
        .Range("E" & RowOut).Value = TreeCBH.CBH
    End With
End Sub

Private Sub Write_AlarmHeader()
    With Worksheets("Alarm")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ObservationID"
        .Range("C1").Value = "WaypointID"
        .Range("D1").Value = "Datim"
        .Range("E1").Value = "PredatorType"
        .Range("F1").Value = "PredatorSpecies"
        .Range("G1").Value = "NumAlarmers"
        .Range("H1").Value = "NumAlarms"
        .Range("I1").Value = "AlarmerAge"
        .Range("J1").Value = "IsConfirmed"
        .Range("K1").Value = "Danger"
        .Range("L1").Value = "IsMultiple"
        .Range("M1").Value = "IsPresent"
        .Range("N1").Value = "ForestLevel"
        .Range("O1").Value = "Height"
    End With
End Sub

Private Sub Write_Alarm(ByRef Alarm As AlarmOutputType)
    Dim RowOut As Long
    RowOut = Alarm.ID + 1
    With Worksheets("Alarm")
        .Range("A" & RowOut).Value = Alarm.ID
        .Range("B" & RowOut).Value = Alarm.ObservationID
        .Range("C" & RowOut).Value = Alarm.WaypointID
        .Range("D" & RowOut).Value = Alarm.Datim
        .Range("E" & RowOut).Value = Alarm.PredatorType
        .Range("F" & RowOut).Value = Alarm.PredatorSpecies
        .Range("G" & RowOut).Value = Alarm.NumAlarmers
        .Range("H" & RowOut).Value = Alarm.NumAlarms
        .Range("I" & RowOut).Value = Alarm.AlarmerAge
        .Range("J" & RowOut).Value = Alarm.IsConfirmed
        .Range("K" & RowOut).Value = Alarm.Danger
        .Range("L" & RowOut).Value = Alarm.IsMultiple
        .Range("M" & RowOut).Value = Alarm.IsPresent
        .Range("N" & RowOut).Value = Alarm.ForestLevel
        If Alarm.Height <> -1 Then .Range("O" & RowOut).Value = Alarm.Height
    End With
End Sub

Private Sub Write_InteractionHeader()
    With Worksheets("Interaction")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ObservationID"
        .Range("C1").Value = "Datim"
        .Range("D1").Value = "Actor"
        .Range("E1").Value = "Recipient"
        .Range("F1").Value = "InteractionType"
    End With
End Sub

Private Sub Write_Interaction(ByRef Interaction As InteractionOutputType)
    Dim RowOut As Long
    RowOut = Interaction.ID + 1
    With Worksheets("Interaction")
        .Range("A" & RowOut).Value = Interaction.ID
        .Range("B" & RowOut).Value = Interaction.ObservationID
        .Range("C" & RowOut).Value = Interaction.Datim
        .Range("D" & RowOut).Value = Interaction.Actor
        .Range("E" & RowOut).Value = Interaction.Recipient
        .Range("F" & RowOut).Value = Interaction.InteractionType
    End With
End Sub

Private Sub Write_IntergroupHeader()
    With Worksheets("Intergroup")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ObservationID"
        .Range("C1").Value = "WaypointID"
        .Range("D1").Value = "Datim"
        .Range("E1").Value = "OpponentGroup"
        .Range("F1").Value = "Outcome"
    End With
End Sub

Private Sub Write_Intergroup(ByRef Intergroup As IntergroupOutputType)
    Dim RowOut As Long
    RowOut = Intergroup.ID + 1
    With Worksheets("Intergroup")
        .Range("A" & RowOut).Value = Intergroup.ID
        .Range("B" & RowOut).Value = Intergroup.ObservationID
        .Range("C" & RowOut).Value = Intergroup.WaypointID
        .Range("D" & RowOut).Value = Intergroup.Datim
        .Range("E" & RowOut).Value = Intergroup.OpponentGroup
        .Range("F" & RowOut).Value = Intergroup.Outcome
    End With
End Sub

Private Sub Write_RangingEventHeader()
    With Worksheets("RangingEvent")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ObservationID"
        .Range("C1").Value = "WaypointID"
        .Range("D1").Value = "Datim"
        .Range("E1").Value = "EventType"
    End With
End Sub

Private Sub Write_RangingEvent(ByRef RangingEvent As RangingEventOutputType)
    Dim RowOut As Long
    RowOut = RangingEvent.ID + 1
    With Worksheets("RangingEvent")
        .Range("A" & RowOut).Value = RangingEvent.ID
        .Range("B" & RowOut).Value = RangingEvent.ObservationID
        .Range("C" & RowOut).Value = RangingEvent.WaypointID
        .Range("D" & RowOut).Value = RangingEvent.Datim
        .Range("E" & RowOut).Value = RangingEvent.EventType
    End With
End Sub

Private Sub Write_CommentHeader()
    With Worksheets("Comment")
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ObservationID"
        .Range("C1").Value = "Datim"
        .Range("D1").Value = "Comment"
    End With
End Sub

Private Sub Write_Comment(ByRef Comment As CommentOutputType)
    Dim RowOut As Long
    RowOut = Comment.ID + 1
    With Worksheets("Comment")
        .Range("A" & RowOut).Value = Comment.ID
        .Range("B" & RowOut).Value = Comment.ObservationID
        .Range("C" & RowOut).Value = Comment.Datim
        .Range("D" & RowOut).Value = Comment.Comment
    End With
End Sub

Private Sub Write_Headers()
    Write_ObservationHeader
    Write_GroupScanHeader
    Write_VertebrateHeader
    Write_FollowHeader
    Write_FollowBlockHeader
    Write_PointSampleHeader
    Write_ActivityHeader
    Write_SelfDirectedHeader
    Write_FoodPatchHeader
    Write_FoodObjectHeader
    Write_ForagingEventHeader
    Write_FruitVisitHeader
    Write_TreeCBHHeader
    Write_AlarmHeader
    Write_InteractionHeader
    Write_IntergroupHeader
    Write_RangingEventHeader
    Write_CommentHeader
End Sub
