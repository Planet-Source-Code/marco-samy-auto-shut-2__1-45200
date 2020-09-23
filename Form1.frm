VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Auto Shutdown"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Force Shut Down (Auto Close Programs)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Tag             =   "Force All Program To Exit And Shut Down, and no program can cancel this shut down... also any unsaved work will be lost!!!!!!!!!!."
      ToolTipText     =   "This Will Close All Programs have data not saved autmatically"
      Top             =   4080
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ask Me Before Shut Down"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Tag             =   "Display a Message Asking You Before Shut Down (Will Not Shut Down Automatically, Or Suddenly!)."
      ToolTipText     =   "If you will sleep, DON'T Check This Box"
      Top             =   3720
      Width           =   3375
   End
   Begin VB.PictureBox Con1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   3585
      TabIndex        =   2
      Top             =   1440
      Width           =   3615
      Begin VB.TextBox tHour 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox tMin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "10"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox tSec 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Automatic Shut Down After :"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hour "
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Minute"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Second"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.PictureBox Con2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   3585
      TabIndex        =   23
      Top             =   1440
      Width           =   3615
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   10  'Mask Pen
         FillColor       =   &H000000FF&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   2505
         TabIndex        =   29
         Top             =   60
         Width           =   2535
      End
      Begin VB.Timer Timer2 
         Interval        =   950
         Left            =   0
         Top             =   1560
      End
      Begin VB.TextBox tSec2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox tMin2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image Lm 
         Height          =   255
         Left            =   720
         MouseIcon       =   "Form1.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":045C
         Stretch         =   -1  'True
         ToolTipText     =   "Chage Usage Value"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   600
         Width           =   495
      End
      Begin VB.Line La 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   120
         X2              =   2880
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Shut Down When Usage is :"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   340
         Width           =   3255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Cur Usage:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Wait By This Usage For :"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Tag             =   $"Form1.frx":0766
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Second"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Minute"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   625
      Top             =   745
      Width           =   3135
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   625
      Top             =   1105
      Width           =   3135
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   1345
      Top             =   3025
      Width           =   2415
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   265
      Top             =   4585
      Width           =   3375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      MouseIcon       =   "Form1.frx":0854
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Tag             =   "is program working now or in the setting mode, click to change (Work Or Pause)."
      ToolTipText     =   "Click Here To Change State"
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Now State :"
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000080&
      X1              =   5640
      X2              =   5640
      Y1              =   5640
      Y2              =   0
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      X1              =   0
      X2              =   5640
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      X1              =   0
      X2              =   0
      Y1              =   360
      Y2              =   5640
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      MouseIcon       =   "Form1.frx":09A6
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Tag             =   "Minimize this window and send it to the taskbar."
      ToolTipText     =   "ÊÕÛíÑ"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quick Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3735
      Left            =   3840
      TabIndex        =   21
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      MouseIcon       =   "Form1.frx":0AF8
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Tag             =   "Display Information about using the program and owners."
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shut Down Now"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Form1.frx":0C4A
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Tag             =   "Do The Shut Down Now Using This Shut Down Options"
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shut Down Options"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Tag             =   "Shut Down Options is some thing program can use in it's shutdown"
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Label Op2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Low CPU Usage"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      MouseIcon       =   "Form1.frx":0D9C
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Tag             =   $"Form1.frx":0EEE
      ToolTipText     =   "Computer Not Working"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Op1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Count Down Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      MouseIcon       =   "Form1.frx":0FCE
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Tag             =   $"Form1.frx":1120
      ToolTipText     =   "Sleep Then It will Shut Down (Check That Ask Me is Removed)"
      Top             =   720
      Width           =   3135
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "OR"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   945
      Width           =   240
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   840
      Y2              =   1200
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Automatic Shut Down Using :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5280
      MouseIcon       =   "Form1.frx":11B1
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Tag             =   "Close This Screen, Direct Exit."
      ToolTipText     =   "ÃÛáÇÞ"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Automatic Shutown Ver 2.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   11
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "By Marco Samy   -       marco_s2@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Tag             =   "Automatic Shutdown Version 2.0, Invented, Desgined and Programmed By Marco Samy - Egypt (El Minia)."
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   100
      Top             =   3000
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In this Project we don't have to put alot of code ...!

'The Only one API function we need to use
'we use it to shutdown windows as you will see
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'Points We Hold to move the screen
'some needed API about time and stop working
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
'Save Points Of Old Mouse Click, We Use It For Moving Form
Dim Ox, Oy
'Indicate Work Or Not
Dim Work As Boolean
'Indicate If we use COUNTDOWN or CPU USAGE
Dim CountDN As Boolean
'Needed Constants
Const NoColor = &H8080FF
Const YesColor = &HC0&
'Class Of Cpu Usage
Dim CPUs As New clsCPUUsage
Dim LowCpu As Long 'Fixed Usage TO Shut down after
Dim LowDone As Boolean
'Some Variable Holding Important Data
Dim OldMinute As Integer, OldSecond As Integer
'/End Delclaration Area

Private Sub Check2_Click()
'this check detremine if the utility is working or not and sets it's value
'as the value of the check2
If Check2.Value = 1 Then Timer1.Enabled = True Else Timer1.Enabled = False
End Sub
Function Timing()
Dim Started As Long 'time point to indicate if second left or not
Started = timeGetTime
Do
If Not Work Then Exit Function 'no we are working or not???
'we must put the old point (Started) before Executing Action, Notice This will make the program work right
If Val(timeGetTime - Started) >= 1000 Then Started = timeGetTime: If CountDN Then EXECDn Else EXECpu Else Sleep 20: DoEvents 'Execute the following function
'do in loop
Loop
End Function
'This Function will be called every second by the previous one
Function EXECDn()
tSec = Val(tSec) - 1
'move to next second
If tSec = -1 Then tSec = 59: tMin = tMin - 1
'move to next minute
If tMin = -1 Then tMin = 59: tHour = tHour - 1
'move to next hour
If tHour = -1 Then Doexit
'check if we are out of time
If Val(tSec) + Val(tMin) + Val(tHour) <= 0 Then Doexit 'now is the time to exit
'change title's caption
Caption = "Shutdown after " & tHour & ":" & tMin & ":" & tSec
End Function
'This Function will be called every second by TIMING
Function EXECpu()
Dim CurCpu As Long
CurCpu = CPUs.CurrentCPUUsage
DrawUsage CurCpu  'drawing new usage
' if already counting down in cpu low usage
If LowDone Then
'check that we still low
If Val(CurCpu) > Val(LowCpu) Then LowDone = False: tMin2 = OldMinute: tSec2 = OldSecond
Else
'if we are not counting down
If Val(CurCpu) <= Val(LowCpu) Then LowDone = True
End If
'Exit Here if not yet low
If Not LowDone Then Exit Function
'But we must put a pass point
tSec2 = Val(tSec2) - 1
'move to next second
If tSec2 = -1 Then tSec2 = 59: tMin2 = tMin2 - 1
'move to next minute
If tMin2 = -1 Then Doexit
'check if we are out of time
If Val(tSec2) + Val(tMin2) <= 0 Then Doexit 'now is the time to exit
'change title's caption
Caption = "Low Cpu Shtdn after " & tMin2 & ":" & tSec2
End Function
'Draw Progress On Picture3
Function DrawUsage(sVal As Long)
Picture3.Cls
'setting line width = all the picture3's width
'Picture3.DrawWidth = (Picture3.Height / Screen.TwipsPerPixelY)
'First Printing Value in the centre
Picture3.CurrentX = (Picture3.Width - Picture3.TextWidth(sVal & "%")) / 2
Picture3.Print sVal & "%"
'Second, Drawing line(s) over it
'For I = 1 To 20 'draw 20 lines
Picture3.Line (0, 0)-(Picture3.Width * (sVal / 100), Picture3.Height), Picture3.FillColor, BF
'Next I
End Function

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Check1.Tag Then Label15.Caption = Check1.Tag
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Check3.Tag Then Label15.Caption = Check3.Tag
End Sub

Private Sub Command1_Click()
'Shutdown now
Doexit 'execute shutdown
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Command1.Tag Then Label15.Caption = Command1.Tag
End Sub

Private Sub Command2_Click()
'about info
MsgBox "Automatic Shut Down Version 2.0" & vbCrLf & "Just Move On Any Item, and It's Description Will Apear On Your Right." & vbCrLf & vbCrLf & "By Marco Samy Nasif" & vbCrLf & vbCrLf & "El-Minia, Egypt.", vbInformation
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Command2.Tag Then Label15.Caption = Command2.Tag
End Sub

Private Sub Form_Load()
'show owr form on the screen
Visible = True
  DoEvents
'By Default Count Down
'Setting Default Values
OldMinute = 1: OldSecond = 0
'
CountDN = True
'Setting Default Value Of CPU Usage in 25%
Label8.Caption = "Auto Shut Down 2.0 - Count Down" 'Changing Title Of View
Vset (25): Lm_MouseMove 1, 1, CSng(Ox), CSng(Oy)  'display new value
'Indicate That we are working
Working (True)
End Sub
Function Working(sVal As Boolean)
'set work value
Work = sVal
If Work Then
'show we are working
Label12.BackColor = vbRed
Label12.ForeColor = vbWhite
Label12.Caption = "Working"
'EXECUTE THE TIMER FUNCTION
Timing
Else
'show not working
Label12.BackColor = NoColor
Label12.ForeColor = vbBlack
Label12.Caption = "Not Working"
End If
End Function

Private Sub Label12_Click()
Working (Not Work) 'did u understant this?
'work indicates current work state
'so we invert it using "NOT" Keyword
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Label15.Caption = Label12.Tag Then Label15.Caption = Label12.Tag
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Label15.Caption = Label14.Tag Then Label15.Caption = Label14.Tag
End Sub

Private Sub Label16_Click()
'minimize the form to the taskbar
WindowState = vbMinimized
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Label16.Tag Then Label15.Caption = Label16.Tag
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Label15.Caption = Label22.Tag Then Label15.Caption = Label22.Tag
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Label6.Tag Then Label15.Caption = Label6.Tag
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Store Old Points
Ox = X: Oy = Y
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Move The Form If the Left Mouse Button is currently being pressed
If Button = 1 Then Move Left + X - Ox, Top + Y - Oy
End Sub

Private Sub Label9_Click()
'exit
End
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Label9.Tag Then Label15.Caption = Label9.Tag
End Sub

Private Sub Op1_Click()
'disabling the shape of op2
Op2.BackColor = NoColor
Op2.ForeColor = vbBlack
Op2.FontBold = False
'enabling op1
Op1.BackColor = YesColor
Op1.ForeColor = vbWhite
Op1.FontBold = True
'Displaying the first option (Count down)
Con1.ZOrder 0
CountDN = True 'count down mode
Label8.Caption = "Auto Shut Down 2.0 - Count Down" 'Changing Title Of View
End Sub

Private Sub Op1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Op1.Tag Then Label15.Caption = Op1.Tag
End Sub

Private Sub Op2_Click()
'disabling the shape of op1
Op1.BackColor = NoColor
Op1.ForeColor = vbBlack
Op1.FontBold = False
'enabling op2
Op2.BackColor = YesColor
Op2.ForeColor = vbWhite
Op2.FontBold = True
'Displaying the second option (CPU)
Con2.ZOrder 0
CountDN = False 'cpu usage mode
Label8.Caption = "Auto Shut Down 2.0 - Lower CPU" 'Changing Title Of View
End Sub

Private Sub Op2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change The Quick Tip Text
If Not Label15.Caption = Op2.Tag Then Label15.Caption = Op2.Tag
End Sub

Private Sub tHour_Change()
On Error Resume Next 'resume if error happened
tHour = Val(tHour) 'Reset Value to numerical value on bad data input
'Maximum Hour Number is 72 no more
If Val(tHour) > 72 Then tHour = 72
If Val(tHour) < -1 Then tHour = 1 'minimal value is 0
End Sub
Function Doexit()
Dim uFlag As Long 'Indicate Which Shutdown Action
Dim RetVal As Long ' value returned by exit windows function
'Selecting Shutdown Action
'Remeber 4 = force shut down, 1 = normal shut down
If CBool(Check3.Value) Then uFlag = 1 Or 4 Else uFlag = 1 'adding force option
'function to execute shudown
Working (False) 'no we are not working(stop counter)
'if user need a question before exit ... then
If Check1.Value = 1 Then If MsgBox("Do you want to shutdown windows now?", vbCritical + vbYesNo) = vbNo Then Exit Function
'the question passed ... we will shutdown now
RetVal = ExitWindowsEx(uFlag, 1)      '1 means normal shutdown,4 means force shutdown
If RetVal = 0 Then 'WinXp need no normal shutdown
'in windows XP(or later) we use shell "shutdown" instead of API function cause ExitWindowsEX always fails
Dim CmdArgs As String 'command line arguments
CmdArgs = "-s" 'shutdown
If CBool(Check3.Value) Then CmdArgs = CmdArgs & " -f" 'force shut down option
CmdArgs = CmdArgs & " -t 01" 'time to wait for shutdown is one second
Shell "Shutdown " & CmdArgs, vbHide   ' execute shut down option
End If 'here are all
End Function

Private Sub tMin_Change()
On Error Resume Next 'resume if error happened
tMin = Val(tMin) 'Reset Value to numerical value on bad data input
'Maximum Minutes Number is 60 no more
If Val(tMin) > 60 Then tMin = 60
If Val(tMin) < -1 Then tMin = 1 'minimal value is 0
End Sub

Private Sub tMin2_Change()
On Error Resume Next 'resume if error happened
tMin2 = Val(tMin2) 'Reset Value to numerical value on bad data input
'Maximum Minutes Number is 60 no more
If Val(tMin2) > 60 Then tMin2 = 60
If Val(tMin2) < -1 Then tMin2 = 1 'minimal value is 0
End Sub

Private Sub tMin2_KeyUp(KeyCode As Integer, Shift As Integer)
'Save Value After Low CPU to Count Down
OldMinute = Abs(Val(tMin2))
OldSecond = Abs(Val(tSec2))
End Sub

Private Sub tSec_Change()
On Error Resume Next 'resume if error happened
tSec = Val(tSec) 'Reset Value to numerical value on bad data input
'Maximum Seconds Number is 60 no more
If Val(tSec) > 60 Then tSec = 60
If Val(tSec) < -1 Then tSec = 1 'minimal value is 0
End Sub

Private Sub tSec2_Change()
On Error Resume Next 'resume if error happened
tSec2 = Val(tSec2) 'Reset Value to numerical value on bad data input
'Maximum Seconds Number is 60 no more
If Val(tSec2) > 60 Then tSec2 = 60
If Val(tSec2) < -1 Then tSec2 = 1 'minimal value is 0
End Sub
'get the value from image position on the line
Function Vget()
Dim iAll As Single
iAll = (La.X2 - La.X1) - Lm.Width
'getting the value from the image's left
Vget = ((Lm.Left - La.X1) / iAll) * 99 + 1
End Function
'return the left of the image position on the line
Function Vset(s_Val)
Dim iAll As Single
iAll = (La.X2 - La.X1) - Lm.Width
'Setting left of the image
Lm.Left = (((s_Val - 1) / 99) * iAll) + La.X1
End Function
Private Sub Lm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set old mouse point
Ox = X
End Sub
'moving image and changing the value
Private Sub Lm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'on button not clicked, exit we need mouse down
If Button <> 1 Then Exit Sub
'skip out of shield
If (Lm.Left + (X - Ox)) < La.X1 Then Lm.Left = La.X1: GoTo Skip
If (Lm.Left + (X - Ox)) > (La.X2 - Lm.Width) Then Lm.Left = La.X2 - Lm.Width: GoTo Skip
'inside shield, so move
Lm.Left = Lm.Left + (X - Ox)
Skip:
'displaying the new value in the next label
LowCpu = Format(Vget, "#")
Label21 = LowCpu & "%"
End Sub

Private Sub tSec2_KeyUp(KeyCode As Integer, Shift As Integer)
'Save Value After Low CPU to Count Down
OldMinute = Abs(Val(tMin2))
OldSecond = Abs(Val(tSec2))
End Sub
'Thank you for downloading this code
'You Can Vote If You Want!
