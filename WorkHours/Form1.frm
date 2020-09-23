VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D6E1E9&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00D6E1E9&
      Height          =   1950
      Left            =   8190
      TabIndex        =   21
      Top             =   510
      Width           =   2130
      Begin VB.ListBox LstHolidays 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         ItemData        =   "Form1.frx":0000
         Left            =   195
         List            =   "Form1.frx":0010
         TabIndex        =   22
         Top             =   570
         Width           =   1680
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Holidays"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   23
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D6E1E9&
      Height          =   1950
      Left            =   10425
      TabIndex        =   15
      Top             =   495
      Width           =   2415
      Begin VB.OptionButton OvertimeIgnore 
         BackColor       =   &H00D6E1E9&
         Caption         =   "Ignore Overtime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         TabIndex        =   17
         Top             =   855
         Value           =   -1  'True
         Width           =   2160
      End
      Begin VB.OptionButton OvertimeInclude 
         BackColor       =   &H00D6E1E9&
         Caption         =   "Include Overtime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   330
         Width           =   2205
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D6E1E9&
      Height          =   1950
      Left            =   5940
      TabIndex        =   11
      Top             =   510
      Width           =   2130
      Begin VB.ListBox List1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         ItemData        =   "Form1.frx":0045
         Left            =   195
         List            =   "Form1.frx":004F
         TabIndex        =   12
         Top             =   570
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Non-Working Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   13
         Top             =   270
         Width           =   1710
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D6E1E9&
      Height          =   1950
      Left            =   150
      TabIndex        =   6
      Top             =   510
      Width           =   3300
      Begin VB.TextBox TxtTEd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Text            =   "29/Feb/2008 6:00 pm"
         Top             =   1350
         Width           =   2820
      End
      Begin VB.TextBox TxtFSt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   195
         TabIndex        =   7
         Text            =   "25/Feb/2008 9:00 am"
         Top             =   510
         Width           =   2820
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date && Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   10
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date && Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   8
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D6E1E9&
      Height          =   1935
      Left            =   3630
      TabIndex        =   1
      Top             =   510
      Width           =   2175
      Begin VB.TextBox TxtWEd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Text            =   "6:00 pm"
         Top             =   1335
         Width           =   1155
      End
      Begin VB.TextBox TxtWSt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Text            =   "9:00 am"
         Top             =   555
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Hr End Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   5
         Top             =   1065
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Hr Start Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   285
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BBCEDC&
      Caption         =   "Calculate Total Working Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2625
      Width           =   3600
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " http://www.nitins.info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   11055
      TabIndex        =   25
      Top             =   270
      Width           =   1785
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For other interesting softwares visit :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   7875
      TabIndex        =   24
      Top             =   270
      Width           =   3135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Total Working Hours Calculation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   19
      Top             =   150
      Width           =   5115
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by : Nitin S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   10845
      TabIndex        =   18
      Top             =   15
      Width           =   1995
   End
   Begin VB.Label Lbltotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   150
      TabIndex        =   14
      Top             =   2685
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Total Working Hours Calculation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   270
      TabIndex        =   20
      Top             =   180
      Width           =   5115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------
'Developed by Nitin S
'For other interesting softwares please visit my personal website www.nitins.info
'---------------------------------------------------------------------------------
Public Function NetWorkhours(dteStart As Date, dteEnd As Date) As Single
    Dim intGrossDays As Integer
    Dim intGrossHours As Single
    Dim dteCurrDate As Date
    Dim i As Integer
    Dim WorkDayStart As Date
    Dim WorkDayend As Date
    Dim nonWorkDays As Integer
    Dim StartDayhours As Single
    Dim EndDayhours As Single
    Dim WorkHrsPerDay As Single
    Dim CntHolidays As Integer
    Dim HolidayHrs As Single
    NetWorkhours = 0
    nonWorkDays = 0
    WorkHrsPerDay = (DateDiff("n", TxtWSt.Text, TxtWEd.Text)) / 60
    
    'Get the work day Start Date/Time and Work day end date/Time
    WorkDayStart = DateValue(dteEnd) + TimeValue(TxtWSt.Text)
    WorkDayend = DateValue(dteStart) + TimeValue(TxtWEd.Text)
      
    'Calculate the Start Day hours
    If IsHolidayOrNonWorkDay(dteStart) = False Then
        StartDayhours = (DateDiff("n", dteStart, WorkDayend)) / 60
    End If
    
    'Calculate the End Day hours
    If IsHolidayOrNonWorkDay(dteEnd) = False Then
        EndDayhours = (DateDiff("n", WorkDayStart, dteEnd)) / 60
    End If
    
    'Nullify Start and End Hours
    If StartDayhours < 0 Then StartDayhours = 0
    If EndDayhours < 0 Then EndDayhours = 0
      
    'Check Overtime hours for Start or End date
    If OvertimeIgnore.Value = True Then
        If StartDayhours > WorkHrsPerDay Then StartDayhours = WorkHrsPerDay
        If EndDayhours > WorkHrsPerDay Then EndDayhours = WorkHrsPerDay
    End If
      
    'Count GrossDay hours and holidays between Gross Days
    If DateDiff("d", DateAdd("d", 1, dteStart), DateAdd("d", -1, dteEnd)) >= 0 Then
         'Calculate total days between the Gross Days
        intGrossDays = DateDiff("d", DateAdd("d", 1, dteStart), DateAdd("d", -1, dteEnd)) + 1
         
         ' Calculate the Gross Hours
        intGrossHours = intGrossDays * WorkHrsPerDay
        
        'Calculate holidays between gross days
        CntHolidays = CountHolidayOrNonWorkDay(DateAdd("d", 1, dteStart), DateAdd("d", -1, dteEnd))
        HolidayHrs = CntHolidays * WorkHrsPerDay
    End If
           
    'Deduct HolidayHours from Gross Hours
    intGrossHours = intGrossHours - HolidayHrs
    
    'Nullify Gross Hours
    If intGrossHours < 0 Then intGrossHours = 0
      
    'Finally Calculate the number of work hours i.e NetWorkhours
    Select Case intGrossDays
        Case 0   'start and end time on same day
            If DateValue(dteStart) = DateValue(dteEnd) Then
                If StartDayhours <> 0 Then NetWorkhours = StartDayhours
            Else
                If StartDayhours <> 0 Then NetWorkhours = StartDayhours
                If EndDayhours <> 0 Then NetWorkhours = NetWorkhours + EndDayhours
            End If
        Case 1   'start and end time on consecutive days
            NetWorkhours = (NetWorkhours + StartDayhours + EndDayhours)
        Case Is > 1  'start and end time on non consecutive days
            NetWorkhours = StartDayhours + intGrossHours + EndDayhours
        End Select
        
    'Nullify NetWorkhours
    If NetWorkhours < 0 Then NetWorkhours = 0
End Function

Private Sub Command1_Click()
    Lbltotal.Caption = "Total : " & NetWorkhours(DateValue(TxtFSt.Text) + TimeValue(TxtFSt.Text), DateValue(TxtTEd.Text) + TimeValue(TxtTEd.Text)) & " Hours"
End Sub
Public Function CountHolidayOrNonWorkDay(StDate, EndDate) As Integer
    Dim Found As Boolean
    For y = StDate To EndDate
        Found = False
        'Check for Holiday
        For x = 0 To LstHolidays.ListCount - 1
            If (DateValue(LstHolidays.List(x)) = DateValue(y)) Then
                Found = True
                
                'Ignore if holiday comes on saturday or sunday
                If Not ((Weekday(LstHolidays.List(x)) = 7) Or (Weekday(LstHolidays.List(x)) = 1)) Then
                    Ctr = Ctr + 1
                End If
            End If
        Next
        
        If Found = False Then
            'Check for Saturday or Sunday
            If ((Weekday(y) = 7) Or (Weekday(y) = 1)) Then
                Ctr = Ctr + 1
            End If
        End If
    Next
    CountHolidayOrNonWorkDay = Ctr
End Function
Public Function IsHolidayOrNonWorkDay(Tmpdate) As Integer
    Dim Found As Boolean
    'Check if  holiday
    For x = 0 To LstHolidays.ListCount - 1
        If (DateValue(LstHolidays.List(x)) = DateValue(Tmpdate)) Then
            Found = True
        End If
    Next
    
    If Found = False Then
        'Check if  saturday or sunday
        If ((Weekday(Tmpdate) = 7) Or (Weekday(Tmpdate) = 1)) Then
            Found = True
        End If
    End If
    
    IsHolidayOrNonWorkDay = Found
End Function

