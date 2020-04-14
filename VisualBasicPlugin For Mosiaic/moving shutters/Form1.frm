VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17385
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   17385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command24 
      Caption         =   "LoadOldImage"
      Height          =   855
      Left            =   15240
      TabIndex        =   67
      Top             =   8040
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   11520
      TabIndex        =   64
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Reset All Make Brighter"
      Height          =   975
      Left            =   12000
      TabIndex        =   63
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reset Single ROI"
      Height          =   615
      Left            =   9960
      TabIndex        =   49
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "TileImage"
      Height          =   735
      Left            =   5280
      TabIndex        =   33
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tiled Image for Nav\Draw"
      Height          =   3135
      Left            =   5040
      TabIndex        =   34
      Top             =   6840
      Width           =   9375
      Begin VB.Frame Frame6 
         Caption         =   "Size"
         Height          =   1815
         Left            =   120
         TabIndex        =   68
         Top             =   1080
         Width           =   1215
         Begin VB.OptionButton Option10 
            Caption         =   "3x3"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1440
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "4x2"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "4x4"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option5 
            Caption         =   "6x6"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option6 
            Caption         =   "8x8"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton Option7 
            Caption         =   "12x12"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Brightness"
         Height          =   2895
         Left            =   6240
         TabIndex        =   62
         Top             =   120
         Width           =   2655
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Larger Change"
            Height          =   615
            Left            =   1800
            TabIndex        =   66
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "No Change"
            Height          =   495
            Left            =   240
            TabIndex        =   65
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Channel(s)"
         Height          =   2775
         Left            =   1680
         TabIndex        =   50
         Top             =   240
         Width           =   2895
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1800
            TabIndex        =   60
            Text            =   "##"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1800
            TabIndex        =   58
            Text            =   "##"
            Top             =   1320
            Width           =   735
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   2160
            Width           =   1095
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton Option9 
            Caption         =   "2"
            Height          =   375
            Left            =   1680
            TabIndex        =   52
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton Option8 
            Caption         =   "1"
            Height          =   495
            Left            =   1680
            TabIndex        =   51
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Exp. Time (ms)"
            Height          =   255
            Left            =   1680
            TabIndex        =   61
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Exp. Time (ms)"
            Height          =   255
            Left            =   1680
            TabIndex        =   59
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Channel 2"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Channel 1"
            Height          =   255
            Left            =   360
            TabIndex        =   56
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "How Many Channels?"
            Height          =   495
            Left            =   360
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Reset All"
         Height          =   735
         Left            =   4920
         TabIndex        =   37
         Top             =   2160
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Draw Mask"
         Height          =   495
         Left            =   4920
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Navigate"
         Height          =   495
         Left            =   4920
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Frame9 
         Caption         =   "Mode?"
         Height          =   1215
         Left            =   4680
         TabIndex        =   38
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   120
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Text            =   "Please Select Position to Fire Mosaic "
      Top             =   0
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Initialize RS232 "
      Height          =   855
      Left            =   15000
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Height          =   5925
      Left            =   5160
      ScaleHeight     =   5865
      ScaleWidth      =   5865
      TabIndex        =   13
      Top             =   480
      Width           =   5925
   End
   Begin VB.Frame Frame5 
      Caption         =   "Stage Control"
      Height          =   6495
      Left            =   2400
      TabIndex        =   9
      Top             =   240
      Width           =   2415
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   240
         TabIndex        =   40
         Text            =   "X"
         Top             =   5280
         Width           =   735
      End
      Begin VB.Frame Frame10 
         Caption         =   "For Debugging Only"
         Height          =   1335
         Left            =   120
         TabIndex        =   39
         Top             =   4800
         Width           =   2175
         Begin VB.TextBox Text6 
            Height          =   495
            Left            =   1200
            TabIndex        =   41
            Text            =   "Y"
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   1320
         TabIndex        =   23
         Text            =   "Y"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Text            =   "X"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Text            =   "Y"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Text            =   "X"
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Select Position"
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   855
      End
      Begin VB.Frame Frame7 
         Caption         =   "Desired xy"
         Height          =   1695
         Left            =   960
         TabIndex        =   14
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Frame Frame15 
         Caption         =   "Tiling Center"
         Height          =   2295
         Left            =   240
         TabIndex        =   78
         Top             =   360
         Width           =   1935
         Begin VB.Label Label16 
            Caption         =   "This is the x and y coordinate that will be used for the center of your tiling"
            Height          =   975
            Left            =   240
            TabIndex        =   79
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Y Center"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "x Center"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   1680
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "End Live"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start Live"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "##"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acquire Image"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image Acquisition"
      Height          =   3735
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      Begin VB.Label Label3 
         Caption         =   "Channel"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Exp. Time (ms)"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Live Controls"
      Height          =   2415
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Mosaic Control"
      Height          =   3135
      Left            =   1560
      TabIndex        =   8
      Top             =   6840
      Width           =   3255
      Begin VB.Frame Frame14 
         Caption         =   "Dichroic?"
         Height          =   2415
         Left            =   1560
         TabIndex        =   75
         Top             =   480
         Width           =   1575
         Begin VB.OptionButton Option13 
            Caption         =   "Split TIRF Cube"
            Height          =   615
            Left            =   240
            TabIndex        =   80
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton Option12 
            Caption         =   "TIRF Cube"
            Height          =   615
            Left            =   240
            TabIndex        =   77
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton Option11 
            Caption         =   "DAPI Cube"
            Height          =   375
            Left            =   240
            TabIndex        =   76
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame12 
         Height          =   2175
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1215
         Begin VB.CommandButton Command18 
            Caption         =   "Fire Mosaic"
            Height          =   735
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Load Mask"
            Height          =   735
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Syringe Controls"
      Height          =   6615
      Left            =   14040
      TabIndex        =   16
      Top             =   240
      Width           =   3015
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   600
         TabIndex        =   48
         Text            =   "#"
         Top             =   5040
         Width           =   495
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Set Rate"
         Height          =   735
         Left            =   1560
         TabIndex        =   47
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Refill"
         Height          =   735
         Left            =   1440
         TabIndex        =   45
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Infuse"
         Height          =   735
         Left            =   360
         TabIndex        =   32
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Set Rate"
         Height          =   735
         Left            =   1560
         TabIndex        =   31
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   600
         TabIndex        =   29
         Text            =   "#"
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Set Target Volume"
         Height          =   615
         Left            =   1560
         TabIndex        =   28
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   600
         TabIndex        =   27
         Text            =   "#"
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Change Syringe Diameter"
         Height          =   735
         Left            =   1440
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Text            =   "#"
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Refill Rate (mL\min)"
         Height          =   375
         Left            =   720
         TabIndex        =   46
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Infuse Rate (mL\min)"
         Height          =   495
         Left            =   600
         TabIndex        =   30
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Target Volume (mL)"
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "SyringeDiameter (mm)"
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   1320
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()

    'User-Selected Illumination Mode
    Dim illMode As String
    
    'getting the current illumination mode
    PubMM.GetMMVariable "Device.Illumination.Setting", illMode
    
    'These are the selections for the illumination combo box
   If Combo1.Text = "GFP 100%" Then
        'reset illumination setting
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
    ElseIf Combo1.Text = "DAPI 100%" Then
        'reset illumination setting
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_dapi_channel.jnl"
    ElseIf Combo1.Text = "Cy5 100%" Then
        'reset illumination setting
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_cy5_channel.jnl"
    ElseIf Combo1.Text = "TxRd 100%" Then
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
    End If
    
  
    
End Sub

Private Sub Command1_Click()

    'Text1.Text = "DoesThisWork"
    'PubMM.LoadImage "D:\Users\JohnE\TIRF Move Oct 2018\100418\561nm excitation\field_488nm_2.tif", 1

    
    'user-defined exposure time
     Dim theExp As String
     theExp = Text2.Text
    
     'Debugging code only
     'Text1.Text = theExp
     
     'converting exposure time to number
     theExpNum = CInt(theExp)
    
    'This is the code to set the exposure time of the camera
    PubMM.SetMMVariable "Camera.Digital.Exposure", theExpNum
    
     'setting the illumination mode
     'PubMM.SetMMVariable "Device.Illumination.Setting", illMode
    
    'This is the code to acquire an image
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"
    
    
End Sub

Private Sub Command10_Click()

'Definitions
Dim NewVolume As Single
Dim NewVolumeString As String
Dim StringSend As String

'Getting the input from the text box
NewVolumeString = Text9.Text

'Making floating point
NewVolume = CSng(NewVolumeString)

'Trying this with a string
StringSend = "DIA " + NewVolumeString

'Changing the Volume
MSComm1.Output = StringSend + Chr(13)


End Sub

Private Sub Command11_Click()

'Definitions
Dim NewVol As Single
Dim NewVolString As String
Dim StringSend1 As String

'Getting the input from the text box
NewVolString = Text10.Text

'Making floating point
NewVol = CSng(NewVolString)

'Trying this with a string
StringSend1 = "TGT " + NewVolString

'Changing the Volume
MSComm1.Output = StringSend1 + Chr(13)

End Sub

'Private Sub Command12_Click()

'This is intialize stage button

'Initial xy positions
'Dim xPosStart As Double
'Dim yPosStart As Double
'Dim xPosStartString As String
'Dim yPosStartString As String

'Getting the initial xy coordinates
'PubMM.GetMMVariable "Device.Stage.XPosition", xPosStart
'PubMM.GetMMVariable "Device.Stage.YPosition", yPosStart

'Making initial xy coordinates strings
'xPosStartString = CStr(xPosStart)
'yPosStartString = CStr(yPosStart)

'Text7.Text = xPosStartString
'Text8.Text = yPosStartString

'End Sub

Private Sub Command13_Click()

'Definitions
Dim InRate As Single
Dim InRateString As String
Dim StringSend2 As String

'Getting the input from the text box
InRateString = Text11.Text

'Making floating point
InRate = CSng(InRateString)

'Trying this with a string
StringSend2 = "RAT " + InRateString + "MM"

'Changing the Volume
MSComm1.Output = StringSend2 + Chr(13)


End Sub

Private Sub Command14_Click()

'This is the button for infusion

'Setting to infuse
Dim InfuseString As String
InfuseString = "DIR INF"
MSComm1.Output = InfuseString + Chr(13)

'Running the Syringe
MSComm1.Output = "RUN" + Chr(13)

End Sub

Private Sub Command15_Click()

'This is the button to try to make a mosaic image

'declarations
Dim xPosMosStart As Double
Dim yPosMosStart As Double
Dim xPosMosStartStr As String
Dim yPosMosStartStr As String
Dim xPosMos(144) As Double
Dim yPosMos(144) As Double
Dim BigImage As Long
Dim BigImage2 As Long
Dim CurrIm As Long
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Pix As Integer
Dim yScale As Integer
Dim xScale As Integer
Dim theWidthMosaic As Integer
Dim XStart As Single
Dim YStart As Single
Dim X As Single
Dim Y As Single
Dim Countx As Integer
Dim CountEightX As Integer
Dim CountTwelveX As Integer


'Initializing the ROI number when I start tiling
ROINum = 0

'Initializing the counter for ROI vertices when I start tiling
MasterROICounter = 0

'initializing the arrays that hold xy coordinates of drawing
For i = 0 To 99
    xDraw(i) = 0.1
    yDraw(i) = 0.1
Next i
                
'initializing all ROI vertices
For r = 0 To 999
    xDrawAllROIs(r) = 0
    yDrawAllROIs(r) = 0
    IdxAllROIs(r) = 0
Next r

'figuring out what color to make the giant images
Dim IsRed1 As Integer
Dim IsRed2 As Integer
IsRed1 = 0
IsRed2 = 0

If Combo2.Text = "TxRd 100%" Then
    IsRed1 = 1
ElseIf Combo2.Text = "TxRd 50%" Then
    IsRed1 = 1
ElseIf Combo2.Text = "TxRd 25%%" Then
    IsRed1 = 1
End If

If Combo3.Text = "TxRd 100%" Then
    IsRed2 = 1
ElseIf Combo3.Text = "TxRd 50%" Then
    IsRed2 = 1
ElseIf Combo3.Text = "TxRd 25%%" Then
    IsRed2 = 1
End If


'Create a big image in which to put tiles - Channel 1
'4x2 mosaic
If Option3.Value = True Then
    PubMM.CreateImage 2048, 1024, 16, "myimage", BigImage
'3x3 mosaic
ElseIf Option10.Value = True Then
    PubMM.CreateImage 1536, 1536, 16, "myimage", BigImage
'4x4 mosaic
ElseIf Option4.Value = True Then
    PubMM.CreateImage 2048, 2048, 16, "myimage", BigImage
'6x6 mosaic
ElseIf Option5.Value = True Then
    PubMM.CreateImage 3072, 3072, 16, "myimage", BigImage
'8x8 mosaic
ElseIf Option6.Value = True Then
    PubMM.CreateImage 4096, 4096, 16, "myimage", BigImage
'12x12 mosaic
ElseIf Option7.Value = True Then
    PubMM.CreateImage 6144, 6144, 16, "myimage", BigImage
End If

'Adjusting color
If IsRed1 = 0 Then
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to make green\jour_make_green.jnl"
Else
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to make red\jour_make_red.jnl"
End If

'Create a big image in which to put tiles - Channel 2
If Option9.Value = True Then
    '4x2 mosaic
    If Option3.Value = True Then
        PubMM.CreateImage 2048, 1024, 16, "myimage2", BigImage2
    '4x4 mosaic
    ElseIf Option4.Value = True Then
        PubMM.CreateImage 2048, 2048, 16, "myimage2", BigImage2
    '6x6 mosaic
    ElseIf Option5.Value = True Then
        PubMM.CreateImage 3072, 3072, 16, "myimage2", BigImage2
    '8x8 mosaic
    ElseIf Option6.Value = True Then
        PubMM.CreateImage 4096, 4096, 16, "myimage2", BigImage2
    '12x12 mosaic
    ElseIf Option7.Value = True Then
        PubMM.CreateImage 6144, 6144, 16, "myimage2", BigImage2
    '3x3 mosaic
    ElseIf Option10.Value = True Then
        PubMM.CreateImage 1536, 1536, 16, "myimage2", BigImage2
    End If
End If

'Adjusting color
If IsRed2 = 0 Then
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to make green\jour_make_green.jnl"
Else
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to make red\jour_make_red.jnl"
End If



'Close the big image
'PubMM.CloseImage BigImage

'get the current stage positions
PubMM.GetMMVariable "Device.Stage.XPosition", xPosMosStart
PubMM.GetMMVariable "Device.Stage.YPosition", yPosMosStart

'Load the current position into text box
xPosMosStartStr = CStr(xPosMosStart)
yPosMosStartStr = CStr(yPosMosStart)
Text7.Text = xPosMosStartStr
Text8.Text = yPosMosStartStr


'''''''''''''''''''''''''''''''''''''''''''''''''''
'making a list of x coordinates

'4x2 mosaic
If Option3.Value = True Then
    xTilePosGlobal(0) = xPosMosStart - 600
    xTilePosGlobal(1) = xPosMosStart - 200
    xTilePosGlobal(2) = xPosMosStart + 200
    xTilePosGlobal(3) = xPosMosStart + 600
    xTilePosGlobal(4) = xPosMosStart - 600
    xTilePosGlobal(5) = xPosMosStart - 200
    xTilePosGlobal(6) = xPosMosStart + 200
    xTilePosGlobal(7) = xPosMosStart + 600
'4x4 mosaic
ElseIf Option4.Value = True Then
    xTilePosGlobal(0) = xPosMosStart - 600
    xTilePosGlobal(1) = xPosMosStart - 200
    xTilePosGlobal(2) = xPosMosStart + 200
    xTilePosGlobal(3) = xPosMosStart + 600
    xTilePosGlobal(4) = xPosMosStart - 600
    xTilePosGlobal(5) = xPosMosStart - 200
    xTilePosGlobal(6) = xPosMosStart + 200
    xTilePosGlobal(7) = xPosMosStart + 600
    xTilePosGlobal(8) = xPosMosStart - 600
    xTilePosGlobal(9) = xPosMosStart - 200
    xTilePosGlobal(10) = xPosMosStart + 200
    xTilePosGlobal(11) = xPosMosStart + 600
    xTilePosGlobal(12) = xPosMosStart - 600
    xTilePosGlobal(13) = xPosMosStart - 200
    xTilePosGlobal(14) = xPosMosStart + 200
    xTilePosGlobal(15) = xPosMosStart + 600
'6x6 mosaic
ElseIf Option5.Value = True Then
    xTilePosGlobal(0) = xPosMosStart - 1000
    xTilePosGlobal(1) = xPosMosStart - 600
    xTilePosGlobal(2) = xPosMosStart - 200
    xTilePosGlobal(3) = xPosMosStart + 200
    xTilePosGlobal(4) = xPosMosStart + 600
    xTilePosGlobal(5) = xPosMosStart + 1000
    
    xTilePosGlobal(6) = xPosMosStart - 1000
    xTilePosGlobal(7) = xPosMosStart - 600
    xTilePosGlobal(8) = xPosMosStart - 200
    xTilePosGlobal(9) = xPosMosStart + 200
    xTilePosGlobal(10) = xPosMosStart + 600
    xTilePosGlobal(11) = xPosMosStart + 1000
    
    xTilePosGlobal(12) = xPosMosStart - 1000
    xTilePosGlobal(13) = xPosMosStart - 600
    xTilePosGlobal(14) = xPosMosStart - 200
    xTilePosGlobal(15) = xPosMosStart + 200
    xTilePosGlobal(16) = xPosMosStart + 600
    xTilePosGlobal(17) = xPosMosStart + 1000
    
    xTilePosGlobal(18) = xPosMosStart - 1000
    xTilePosGlobal(19) = xPosMosStart - 600
    xTilePosGlobal(20) = xPosMosStart - 200
    xTilePosGlobal(21) = xPosMosStart + 200
    xTilePosGlobal(22) = xPosMosStart + 600
    xTilePosGlobal(23) = xPosMosStart + 1000
    
    xTilePosGlobal(24) = xPosMosStart - 1000
    xTilePosGlobal(25) = xPosMosStart - 600
    xTilePosGlobal(26) = xPosMosStart - 200
    xTilePosGlobal(27) = xPosMosStart + 200
    xTilePosGlobal(28) = xPosMosStart + 600
    xTilePosGlobal(29) = xPosMosStart + 1000
    
    xTilePosGlobal(30) = xPosMosStart - 1000
    xTilePosGlobal(31) = xPosMosStart - 600
    xTilePosGlobal(32) = xPosMosStart - 200
    xTilePosGlobal(33) = xPosMosStart + 200
    xTilePosGlobal(34) = xPosMosStart + 600
    xTilePosGlobal(35) = xPosMosStart + 1000
    
'8x8 mosaic
ElseIf Option6.Value = True Then

    'initialize counter
    CountEightX = 0

    'decided to loop here
    For v = 0 To 63
        
        If CountEightX = 0 Then
            xTilePosGlobal(v) = xPosMosStart - 1400
        ElseIf CountEightX = 1 Then
            xTilePosGlobal(v) = xPosMosStart - 1000
        ElseIf CountEightX = 2 Then
            xTilePosGlobal(v) = xPosMosStart - 600
        ElseIf CountEightX = 3 Then
            xTilePosGlobal(v) = xPosMosStart - 200
        ElseIf CountEightX = 4 Then
            xTilePosGlobal(v) = xPosMosStart + 200
        ElseIf CountEightX = 5 Then
            xTilePosGlobal(v) = xPosMosStart + 600
        ElseIf CountEightX = 6 Then
            xTilePosGlobal(v) = xPosMosStart + 1000
        ElseIf CountEightX = 7 Then
            xTilePosGlobal(v) = xPosMosStart + 1400
        End If
        
        'iterate counter
        CountEightX = CountEightX + 1
        
        'resetting counter
        If CountEightX = 8 Then
            CountEightX = 0
        End If
        
    Next v
    
'12x12 mosaic
ElseIf Option7.Value = True Then

    'initialize counter
    CountTwelveX = 0

    'decided to loop here
    For w = 0 To 143
        
        If CountTwelveX = 0 Then
            xTilePosGlobal(w) = xPosMosStart - 2200
        ElseIf CountTwelveX = 1 Then
            xTilePosGlobal(w) = xPosMosStart - 1800
        ElseIf CountTwelveX = 2 Then
            xTilePosGlobal(w) = xPosMosStart - 1400
        ElseIf CountTwelveX = 3 Then
            xTilePosGlobal(w) = xPosMosStart - 1000
        ElseIf CountTwelveX = 4 Then
            xTilePosGlobal(w) = xPosMosStart - 600
        ElseIf CountTwelveX = 5 Then
            xTilePosGlobal(w) = xPosMosStart - 200
        ElseIf CountTwelveX = 6 Then
            xTilePosGlobal(w) = xPosMosStart + 200
        ElseIf CountTwelveX = 7 Then
            xTilePosGlobal(w) = xPosMosStart + 600
        ElseIf CountTwelveX = 8 Then
            xTilePosGlobal(w) = xPosMosStart + 1000
        ElseIf CountTwelveX = 9 Then
            xTilePosGlobal(w) = xPosMosStart + 1400
        ElseIf CountTwelveX = 10 Then
            xTilePosGlobal(w) = xPosMosStart + 1800
        ElseIf CountTwelveX = 11 Then
            xTilePosGlobal(w) = xPosMosStart + 2200
        End If
        
        'iterate counter
        CountTwelveX = CountTwelveX + 1
        
        'resetting counter
        If CountTwelveX = 12 Then
            CountTwelveX = 0
        End If
        
    Next w
    
'3x3 mosaic
ElseIf Option10.Value = True Then
    xTilePosGlobal(0) = xPosMosStart - 400
    xTilePosGlobal(1) = xPosMosStart
    xTilePosGlobal(2) = xPosMosStart + 400
    xTilePosGlobal(3) = xPosMosStart - 400
    xTilePosGlobal(4) = xPosMosStart
    xTilePosGlobal(5) = xPosMosStart + 400
    xTilePosGlobal(6) = xPosMosStart - 400
    xTilePosGlobal(7) = xPosMosStart
    xTilePosGlobal(8) = xPosMosStart + 400
End If


''''''''''''''''''''''''''''''''
'making a list of y coordinates

'4x2 mosaic
If Option3.Value = True Then
    yTilePosGlobal(0) = yPosMosStart + 200
    yTilePosGlobal(1) = yPosMosStart + 200
    yTilePosGlobal(2) = yPosMosStart + 200
    yTilePosGlobal(3) = yPosMosStart + 200
    yTilePosGlobal(4) = yPosMosStart - 200
    yTilePosGlobal(5) = yPosMosStart - 200
    yTilePosGlobal(6) = yPosMosStart - 200
    yTilePosGlobal(7) = yPosMosStart - 200
'4x4 mosaic
ElseIf Option4.Value = True Then
    yTilePosGlobal(0) = yPosMosStart + 600
    yTilePosGlobal(1) = yPosMosStart + 600
    yTilePosGlobal(2) = yPosMosStart + 600
    yTilePosGlobal(3) = yPosMosStart + 600
    yTilePosGlobal(4) = yPosMosStart + 200
    yTilePosGlobal(5) = yPosMosStart + 200
    yTilePosGlobal(6) = yPosMosStart + 200
    yTilePosGlobal(7) = yPosMosStart + 200
    yTilePosGlobal(8) = yPosMosStart - 200
    yTilePosGlobal(9) = yPosMosStart - 200
    yTilePosGlobal(10) = yPosMosStart - 200
    yTilePosGlobal(11) = yPosMosStart - 200
    yTilePosGlobal(12) = yPosMosStart - 600
    yTilePosGlobal(13) = yPosMosStart - 600
    yTilePosGlobal(14) = yPosMosStart - 600
    yTilePosGlobal(15) = yPosMosStart - 600
'6x6 mosaic
ElseIf Option5.Value = True Then
    yTilePosGlobal(0) = yPosMosStart + 1000
    yTilePosGlobal(1) = yPosMosStart + 1000
    yTilePosGlobal(2) = yPosMosStart + 1000
    yTilePosGlobal(3) = yPosMosStart + 1000
    yTilePosGlobal(4) = yPosMosStart + 1000
    yTilePosGlobal(5) = yPosMosStart + 1000
    
    yTilePosGlobal(6) = yPosMosStart + 600
    yTilePosGlobal(7) = yPosMosStart + 600
    yTilePosGlobal(8) = yPosMosStart + 600
    yTilePosGlobal(9) = yPosMosStart + 600
    yTilePosGlobal(10) = yPosMosStart + 600
    yTilePosGlobal(11) = yPosMosStart + 600
    
    yTilePosGlobal(12) = yPosMosStart + 200
    yTilePosGlobal(13) = yPosMosStart + 200
    yTilePosGlobal(14) = yPosMosStart + 200
    yTilePosGlobal(15) = yPosMosStart + 200
    yTilePosGlobal(16) = yPosMosStart + 200
    yTilePosGlobal(17) = yPosMosStart + 200

    yTilePosGlobal(18) = yPosMosStart - 200
    yTilePosGlobal(19) = yPosMosStart - 200
    yTilePosGlobal(20) = yPosMosStart - 200
    yTilePosGlobal(21) = yPosMosStart - 200
    yTilePosGlobal(22) = yPosMosStart - 200
    yTilePosGlobal(23) = yPosMosStart - 200
    
    yTilePosGlobal(24) = yPosMosStart - 600
    yTilePosGlobal(25) = yPosMosStart - 600
    yTilePosGlobal(26) = yPosMosStart - 600
    yTilePosGlobal(27) = yPosMosStart - 600
    yTilePosGlobal(28) = yPosMosStart - 600
    yTilePosGlobal(29) = yPosMosStart - 600
    
    yTilePosGlobal(30) = yPosMosStart - 1000
    yTilePosGlobal(31) = yPosMosStart - 1000
    yTilePosGlobal(32) = yPosMosStart - 1000
    yTilePosGlobal(33) = yPosMosStart - 1000
    yTilePosGlobal(34) = yPosMosStart - 1000
    yTilePosGlobal(35) = yPosMosStart - 1000
    
'8x8 mosaic
ElseIf Option6.Value = True Then

    'decided to loop here
    For w = 0 To 63
        
        If w < 8 Then
            yTilePosGlobal(w) = yPosMosStart + 1400
        ElseIf w > 7 And w < 16 Then
            yTilePosGlobal(w) = yPosMosStart + 1000
        ElseIf w > 15 And w < 24 Then
            yTilePosGlobal(w) = yPosMosStart + 600
        ElseIf w > 23 And w < 32 Then
            yTilePosGlobal(w) = yPosMosStart + 200
        ElseIf w > 31 And w < 40 Then
            yTilePosGlobal(w) = yPosMosStart - 200
        ElseIf w > 39 And w < 48 Then
            yTilePosGlobal(w) = yPosMosStart - 600
        ElseIf w > 47 And w < 56 Then
            yTilePosGlobal(w) = yPosMosStart - 1000
        ElseIf w > 55 Then
            yTilePosGlobal(w) = yPosMosStart - 1400
        End If
    
    Next w

'12x12 tiled image
ElseIf Option7.Value = True Then

    'decided to loop here
    For s = 0 To 143
        
        If s < 12 Then
            yTilePosGlobal(s) = yPosMosStart + 2200
        ElseIf s > 11 And s < 24 Then
            yTilePosGlobal(s) = yPosMosStart + 1800
        ElseIf s > 23 And s < 36 Then
            yTilePosGlobal(s) = yPosMosStart + 1400
        ElseIf s > 35 And s < 48 Then
            yTilePosGlobal(s) = yPosMosStart + 1000
        ElseIf s > 47 And s < 60 Then
            yTilePosGlobal(s) = yPosMosStart + 600
        ElseIf s > 59 And s < 72 Then
            yTilePosGlobal(s) = yPosMosStart + 200
        ElseIf s > 71 And s < 84 Then
            yTilePosGlobal(s) = yPosMosStart - 200
        ElseIf s > 83 And s < 96 Then
            yTilePosGlobal(s) = yPosMosStart - 600
        ElseIf s > 95 And s < 108 Then
            yTilePosGlobal(s) = yPosMosStart - 1000
        ElseIf s > 107 And s < 120 Then
            yTilePosGlobal(s) = yPosMosStart - 1400
        ElseIf s > 119 And s < 132 Then
            yTilePosGlobal(s) = yPosMosStart - 1800
        ElseIf s > 131 And s < 144 Then
            yTilePosGlobal(s) = yPosMosStart - 2200
        End If
    
    Next s
    
'3x3 tiled image
ElseIf Option10.Value = True Then
        yTilePosGlobal(0) = yPosMosStart + 400
        yTilePosGlobal(1) = yPosMosStart + 400
        yTilePosGlobal(2) = yPosMosStart + 400
        yTilePosGlobal(3) = yPosMosStart
        yTilePosGlobal(4) = yPosMosStart
        yTilePosGlobal(5) = yPosMosStart
        yTilePosGlobal(6) = yPosMosStart - 400
        yTilePosGlobal(7) = yPosMosStart - 400
        yTilePosGlobal(8) = yPosMosStart - 400
    
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Acquiring and saving images

'illumination setting
Dim LightSetting As String
Dim IsItRed As Integer

'initializing red flag
IsItRed = 0


'user-defined exposure time
Dim theExpCh1 As String
Dim theExpCh1Num As Integer
Dim theExpCh2 As String
Dim theExpCh2Num As Integer

If Option8.Value = True Then
    theExpCh1 = Text13.Text
    theExpCh1Num = CInt(theExpCh1)
End If

If Option9.Value = True Then
    theExpCh1 = Text13.Text
    theExpCh1Num = CInt(theExpCh1)
    theExpCh2 = Text14.Text
    theExpCh2Num = CInt(theExpCh2)
End If


'Dim theExp As String
'Dim theExpNum As Integer
'theExp = Text2.Text
'theExpNum = CInt(theExp)

'This is the code to set the exposure time of the camera
'PubMM.SetMMVariable "Camera.Digital.Exposure", theExpNum

'4x2 mosaic - making big image
If Option3.Value = True Then

    For i = 0 To 7
    
        'making picture box asymmetic
        Picture1.Width = 8000
        Picture1.Height = 4000
        Picture1.Left = 5160
        Picture1.Top = 720

        'making the mosaic
        If i = 0 Then
            xScale = 0
            yScale = 0
        ElseIf i = 1 Then
            xScale = 512
            yScale = 0
        ElseIf i = 2 Then
            xScale = 1024
            yScale = 0
        ElseIf i = 3 Then
            xScale = 1536
            yScale = 0
        ElseIf i = 4 Then
            xScale = 0
            yScale = 512
        ElseIf i = 5 Then
            xScale = 512
            yScale = 512
        ElseIf i = 6 Then
            xScale = 1024
            yScale = 512
        ElseIf i = 7 Then
            xScale = 1536
            yScale = 512
        End If
        
        'Move the stage
        PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(i)
        PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(i)
        
        'Acquiring Channel 1
        
            If Option8.Value = True Or Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh1Num
            
                'Setting the channel to channel 1
                If Combo2.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo2.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo2.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                ElseIf Combo2.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 25%%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                    IsItRed = 1
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3
                

                'Close the image
                PubMM.CloseImage CurrIm
            
            End If
            
        'Acquiring Channel 2

            If Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh2Num
            
                'Setting the channel to channel 1
                If Combo3.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                ElseIf Combo3.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                ElseIf Combo3.Text = "TxRd 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                ElseIf Combo3.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo3.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo3.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage2, j + xScale, k + yScale, Pix
                Next k
                Next j
    
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3
    
                'Close the image
                PubMM.CloseImage CurrIm
            
            End If

    Next i
    
    '4x4 mosaic
    ElseIf Option4.Value = True Then
    
        'making picture box symmetic
        Picture1.Width = 5925
        Picture1.Height = 5925
        Picture1.Left = 6360
        Picture1.Top = 720
    
        'Some Counter
        Countx = 0

        For i = 0 To 15
        
            'xScale
            xScale = Countx * 512
            
            'x counter
            Countx = Countx + 1
            If Countx = 4 Then
                Countx = 0
            End If

            'yScale
            If i < 4 Then
                yScale = 0
            ElseIf i > 3 And i < 8 Then
                yScale = 512
            ElseIf i > 7 And i < 12 Then
                yScale = 1024
            ElseIf i > 11 Then
                yScale = 1536
            End If
        
            'Move the stage
            PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(i)
            PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(i)

            'Acquiring Channel 1
        
            If Option8.Value = True Or Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh1Num
            
                'Setting the channel to channel 1
                If Combo2.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo2.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo2.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                ElseIf Combo2.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 25%%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                    IsItRed = 1
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3

                'Close the image
                PubMM.CloseImage CurrIm
            
            End If
            
        'Acquiring Channel 2

            If Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh2Num
            
                'Setting the channel to channel 1
                If Combo3.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                ElseIf Combo3.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                ElseIf Combo3.Text = "TxRd 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                ElseIf Combo3.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo3.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo3.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage2, j + xScale, k + yScale, Pix
                Next k
                Next j
    
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3
    
                'Close the image
                PubMM.CloseImage CurrIm
            
            End If

        Next i
    
    '6x6 tiled image
    ElseIf Option5.Value = True Then
    
    'intialize counter
    Countx = 0
    
    '6x6 mosaic
    For i = 0 To 35
    
            'making picture box symmetic
            Picture1.Width = 5925
            Picture1.Height = 5925
            Picture1.Left = 6360
            Picture1.Top = 720

            'xScale
            xScale = Countx * 512
            
            'x counter
            Countx = Countx + 1
            If Countx = 6 Then
                Countx = 0
            End If
            
            'yScale
            If i < 6 Then
                yScale = 0
            ElseIf i > 5 And i < 12 Then
                yScale = 512
            ElseIf i > 11 And i < 18 Then
                yScale = 1024
            ElseIf i > 17 And i < 24 Then
                yScale = 1536
            ElseIf i > 23 And i < 30 Then
                yScale = 2048
            ElseIf i > 29 And i < 36 Then
                yScale = 2560
            End If
            
            'Move the stage
             PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(i)
             PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(i)

            'Acquiring Channel 1
            If Option8.Value = True Or Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh1Num
            
                'Setting the channel to channel 1
                If Combo2.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo2.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo2.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                ElseIf Combo2.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 25%%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                    IsItRed = 1
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3

                'Close the image
                PubMM.CloseImage CurrIm
            
            End If
            
        'Acquiring Channel 2

            If Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh2Num
            
                'Setting the channel to channel 1
                If Combo3.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                ElseIf Combo3.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                ElseIf Combo3.Text = "TxRd 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                ElseIf Combo3.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo3.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo3.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage2, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3
    
                'Close the image
                PubMM.CloseImage CurrIm
            
            End If

            
            
        Next i
    
    '8x8 tiled image
    ElseIf Option6.Value = True Then
    
    'intialize counter
    Countx = 0
    
    '8x8 mosaic
    For i = 0 To 63
    
            'making picture box symmetic
            Picture1.Width = 5925
            Picture1.Height = 5925
            Picture1.Left = 6360
            Picture1.Top = 720

            'xScale
            xScale = Countx * 512
            
            'x counter
            Countx = Countx + 1
            If Countx = 8 Then
                Countx = 0
            End If
            
            'yScale
            If i < 8 Then
                yScale = 0
            ElseIf i > 7 And i < 16 Then
                yScale = 512
            ElseIf i > 15 And i < 24 Then
                yScale = 1024
            ElseIf i > 23 And i < 32 Then
                yScale = 1536
            ElseIf i > 31 And i < 40 Then
                yScale = 2048
            ElseIf i > 39 And i < 48 Then
                yScale = 2560
             ElseIf i > 47 And i < 56 Then
                yScale = 3072
            ElseIf i > 55 Then
                yScale = 3584
            End If
            
            'Move the stage
             PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(i)
             PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(i)
             
              'Acquiring Channel 1
            If Option8.Value = True Or Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh1Num
            
                'Setting the channel to channel 1
                If Combo2.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo2.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo2.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                ElseIf Combo2.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 25%%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                    IsItRed = 1
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3

                'Close the image
                PubMM.CloseImage CurrIm
            
            End If
            
        'Acquiring Channel 2

            If Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh2Num
            
                'Setting the channel to channel 1
                If Combo3.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                ElseIf Combo3.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                ElseIf Combo3.Text = "TxRd 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                ElseIf Combo3.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo3.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo3.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage2, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3
    
                'Close the image
                PubMM.CloseImage CurrIm
            
            End If

             
    Next i
    
    '12x12 tiled image
    ElseIf Option7.Value = True Then
 
    
    'intialize counter
    Countx = 0
    
    '12x12 mosaic
    For i = 0 To 143
    
            'making picture box symmetic
            Picture1.Width = 5925
            Picture1.Height = 5925
            Picture1.Left = 6360
            Picture1.Top = 720
    
            
            'xScale
            xScale = Countx * 512
            
            'x counter
            Countx = Countx + 1
            If Countx = 12 Then
                Countx = 0
            End If
            
            'yScale
            If i < 12 Then
                yScale = 0
            ElseIf i > 11 And i < 24 Then
                yScale = 512
            ElseIf i > 23 And i < 36 Then
                yScale = 1024
            ElseIf i > 35 And i < 48 Then
                yScale = 1536
            ElseIf i > 47 And i < 60 Then
                yScale = 2048
            ElseIf i > 59 And i < 72 Then
                yScale = 2560
             ElseIf i > 71 And i < 84 Then
                yScale = 3072
            ElseIf i > 83 And i < 96 Then
                yScale = 3584
            ElseIf i > 95 And i < 108 Then
                yScale = 4096
            ElseIf i > 107 And i < 120 Then
                yScale = 4608
            ElseIf i > 119 And i < 132 Then
                yScale = 5120
            ElseIf i > 131 And i < 144 Then
                yScale = 5632
            End If
            
            'Move the stage
             PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(i)
             PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(i)
             
             
               'Acquiring Channel 1
            If Option8.Value = True Or Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh1Num
            
                'Setting the channel to channel 1
                If Combo2.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo2.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo2.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                ElseIf Combo2.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 25%%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                    IsItRed = 1
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3

                'Close the image
                PubMM.CloseImage CurrIm
            
            End If
            
        'Acquiring Channel 2

            If Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh2Num
            
                'Setting the channel to channel 1
                If Combo3.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                ElseIf Combo3.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                ElseIf Combo3.Text = "TxRd 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                ElseIf Combo3.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo3.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo3.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage2, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3
    
                'Close the image
                PubMM.CloseImage CurrIm
            
            End If

            
    Next i

    '3x3 mosaic
    ElseIf Option10.Value = True Then
    
        'making picture box symmetic
        Picture1.Width = 5925
        Picture1.Height = 5925
        Picture1.Left = 6360
        Picture1.Top = 720
    
        'Some Counter
        Countx = 0

        For i = 0 To 8
        
            'xScale
            xScale = Countx * 512
            
            'x counter
            Countx = Countx + 1
            If Countx = 3 Then
                Countx = 0
            End If

            'yScale
            If i < 3 Then
                yScale = 0
            ElseIf i > 2 And i < 6 Then
                yScale = 512
            ElseIf i > 5 And i < 9 Then
                yScale = 1024
            End If
        
            'Move the stage
            PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(i)
            PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(i)

            'Acquiring Channel 1
        
            If Option8.Value = True Or Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh1Num
            
                'Setting the channel to channel 1
                If Combo2.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo2.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo2.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                ElseIf Combo2.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                    IsItRed = 1
                ElseIf Combo2.Text = "TxRd 25%%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                    IsItRed = 1
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage, j + xScale, k + yScale, Pix
                Next k
                Next j
                
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3

                'Close the image
                PubMM.CloseImage CurrIm
            
            End If
            
        'Acquiring Channel 2

            If Option9.Value = True Then

                'set the exposure time
                PubMM.SetMMVariable "Camera.Digital.Exposure", theExpCh2Num
            
                'Setting the channel to channel 1
                If Combo3.Text = "TxRd 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
                ElseIf Combo3.Text = "TxRd 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_50.jnl"
                ElseIf Combo3.Text = "TxRd 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_25.jnl"
                ElseIf Combo3.Text = "GFP 100%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
                ElseIf Combo3.Text = "GFP 50%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_50.jnl"
                ElseIf Combo3.Text = "GFP 25%" Then
                    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel_25.jnl"
                End If
    
                'This is the code to acquire an image
                PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to acquire image gfp\gfp_acquire.jnl"

                'Get the current image
                PubMM.GetCurrentImage CurrIm
    
                For j = 0 To 511
                For k = 0 To 511
                    PubMM.ReadPixel CurrIm, j, k, Pix
                    PubMM.WritePixel BigImage2, j + xScale, k + yScale, Pix
                Next k
                Next j
    
                'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
                PubMM.SaveImage CurrIm, "D:\Users\JohnE\tmp\tmpImage.tif", False, 3
    
                'Close the image
                PubMM.CloseImage CurrIm
            
            End If

        Next i
    
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Saving Image to put into Picture Box'''''''''''''''''''''

'initialization
Dim FileExt As Integer
Dim FileExt2 As Integer
Dim bright As Double
Dim CurrPlane As Integer
FileExt = 2
FileExt2 = 3

'Single Channel
If Option9.Value = False Then

    'Changing Color if necessary
    If IsItRed = 1 Then
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to make red\jour_make_red.jnl"
    Else
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to make green\jour_make_green.jnl"
    End If
    
    'Get the brightness and hold onto it
    PubMM.GetBrightness BigImage, bright
    TheBrightness = bright
    
    
    'This is some new code to see if I can add a header to this image
    
    'Dim HeaderTest(100) As String
    'HeaderTest(0) = "Line 1"
    'HeaderTest(1) = "Line 2"
    'HeaderTest(2) = "Line 3"
    'PubMM.SetImageAnnotation BigImage, CurrPlane, "Hello", "Hi"
    
    'Trying to output the annotation
    'Dim Anno As String
    'PubMM.GetImageAnnotation BigImage, CurrPlane, Anno
    'MsgBox Anno

    'Saving the big tiled image as *.bmp
    PubMM.SaveImage BigImage, "D:\Users\JohnE\VisualBasicTests\TileTest2.bmp", False, FileExt

    'putting a picture in gui
    Picture1.Picture = LoadPicture("D:\Users\JohnE\VisualBasicTests\TileTest2.bmp")
    
    'forcing the image to fit in the box
    Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    
    'displaying the illumination setting
    'PubMM.GetMMVariable "Device.Illumination.Setting", LightSetting
    'Text4.Text = LightSetting
    
    'Saving image as *.tif
    PubMM.SaveImage BigImage, "D:\Users\JohnE\VisualBasicTests\Ch1_TiledImageAsTif.tif", False, FileExt2

    'close the big image
    PubMM.CloseImage BigImage

'2 channel
ElseIf Option9.Value = True Then

    'save the two images as *.tif first
    PubMM.SaveImage BigImage, "D:\Users\JohnE\VisualBasicTests\Ch1_TiledImageAsTif.tif", False, FileExt2
    PubMM.SaveImage BigImage2, "D:\Users\JohnE\VisualBasicTests\Ch2_TiledImageAsTif.tif", False, FileExt2

    'Merge the two images in color space
    If IsItRed = 0 Then
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to merge images\jour_color_combine.jnl"
    Else
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to merge images\jour_color_combine_rev_order.jnl"
    End If
    
    'Debugging
    'Text4.Text = CStr(IsItRed)
    
    'get the merged image
    PubMM.GetCurrentImage CurrIm
    
    'Get the brightness and hold onto it
    PubMM.GetBrightness CurrIm, bright
    TheBrightness = bright
    
    'Saving the big tiled image as *.bmp
    PubMM.SaveImage CurrIm, "D:\Users\JohnE\VisualBasicTests\TileTest2.bmp", False, FileExt

    'putting a picture in gui
    Picture1.Picture = LoadPicture("D:\Users\JohnE\VisualBasicTests\TileTest2.bmp")

    'forcing the image to fit in the box
    Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    
    'close the big image
    PubMM.CloseImage CurrIm

End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''This is new code which will output text file in order''''''''''''''''''''''''
'''''''''''''''''''''''''''for masks to be reloaded''''''''''''''''''''''''''''''''''''''''''

'Open "D:\Users\JohnE\VisualBasicTests\OutPutTest.txt" For Output As #1
'Print #1, "Does this print to the file?"
'Print #1, "Is this on second line?"
'Print #1, "Is this on third line?"
'Print #1, "Is this on fourth line?"
'Print #1, "Is this on fifth line?"
'Close #1

'definitions
Dim SizeForTxtFileStr As String
Dim LoopNumForTxtFile As Integer
Dim Picture1WidthTxt As String
Dim Picture1HeightTxt As String
Dim Picture1LeftTxt As String
Dim Picture1TopTxt As String

'Figuring out the size of the tiled image
If Option3.Value = True Then
    '4x2 tiled image
    SizeForTxtFileStr = CStr(2)
    LoopNumForTxtFile = 7
ElseIf Option4.Value = True Then
    '4x4 tiled image
    SizeForTxtFileStr = CStr(4)
    LoopNumForTxtFile = 15
ElseIf Option5.Value = True Then
    '6x6 tiled image
    SizeForTxtFileStr = CStr(6)
    LoopNumForTxtFile = 35
ElseIf Option6.Value = True Then
    '8x8 tiled image
    SizeForTxtFileStr = CStr(8)
    LoopNumForTxtFile = 63
ElseIf Option7.Value = True Then
    '12x12 tiled image
    SizeForTxtFileStr = CStr(12)
    LoopNumForTxtFile = 143
ElseIf Option10.Value = True Then
    '3x3 tiled image
    SizeForTxtFileStr = CStr(3)
    LoopNumForTxtFile = 8
End If

'open the text file to write to
Open "D:\Users\JohnE\VisualBasicTests\OutPutTest.txt" For Output As #1

'Loading the size and coordinate centers to text file
Print #1, SizeForTxtFileStr
Print #1, Text7.Text
Print #1, Text8.Text

'Loading the information for the PictureBox
Picture1WidthTxt = CStr(Picture1.Width)
Picture1HeightTxt = CStr(Picture1.Height)
Picture1LeftTxt = CStr(Picture1.Left)
Picture1TopTxt = CStr(Picture1.Top)
Print #1, Picture1WidthTxt
Print #1, Picture1HeightTxt
Print #1, Picture1LeftTxt
Print #1, Picture1TopTxt

'loading the tile coordinates into text file
For v = 0 To LoopNumForTxtFile
    Print #1, CStr(xTilePosGlobal(v))
    Print #1, CStr(yTilePosGlobal(v))
Next v

'closing the text file
Close #1


End Sub

Private Sub Command16_Click()
    
        'Reset ROI number
        ROINum = 0
        
        'putting a picture in gui
        Picture1.Picture = LoadPicture("D:\Users\JohnE\VisualBasicTests\TileTest2.bmp")

        'forcing the image to fit in the box
        Picture1.ScaleMode = 3
        Picture1.AutoRedraw = True
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
        
        'initializing the counter in which I keep track of drawing of coordinates
        MasterROICounter = 0
    
        'initializing the arrays that hold xy coordinates of drawing
        For i = 0 To 99
            xDraw(i) = 0.1
            yDraw(i) = 0.1
        Next i
                
        'initializing all ROI vertices
        For r = 0 To 999
            xDrawAllROIs(r) = 0
            yDrawAllROIs(r) = 0
            IdxAllROIs(r) = 0
         Next r


End Sub

Private Sub Command17_Click()
'This is the button to the load the big (tiled) mask to use for mosaic firing
Dim BigMask As Long
Dim Tile0 As Long
Dim Tile1 As Long
Dim Tile2 As Long
Dim Tile3 As Long
Dim Tile4 As Long
Dim Tile5 As Long
Dim Tile6 As Long
Dim Tile7 As Long
Dim Tile8 As Long
Dim Tile9 As Long
Dim Tile10 As Long
Dim Tile11 As Long
Dim Tile12 As Long
Dim Tile13 As Long
Dim Tile14 As Long
Dim Tile15 As Long
Dim Tile16 As Long
Dim Tile17 As Long
Dim Tile18 As Long
Dim Tile19 As Long
Dim Tile20 As Long
Dim Tile21 As Long
Dim Tile22 As Long
Dim Tile23 As Long
Dim Tile24 As Long
Dim Tile25 As Long
Dim Tile26 As Long
Dim Tile27 As Long
Dim Tile28 As Long
Dim Tile29 As Long
Dim Tile30 As Long
Dim Tile31 As Long
Dim Tile32 As Long
Dim Tile33 As Long
Dim Tile34 As Long
Dim Tile35 As Long
Dim Tile36 As Long
Dim Tile37 As Long
Dim Tile38 As Long
Dim Tile39 As Long
Dim Tile40 As Long
Dim Tile41 As Long
Dim Tile42 As Long
Dim Tile43 As Long
Dim Tile44 As Long
Dim Tile45 As Long
Dim Tile46 As Long
Dim Tile47 As Long
Dim Tile48 As Long
Dim Tile49 As Long
Dim Tile50 As Long
Dim Tile51 As Long
Dim Tile52 As Long
Dim Tile53 As Long
Dim Tile54 As Long
Dim Tile55 As Long
Dim Tile56 As Long
Dim Tile57 As Long
Dim Tile58 As Long
Dim Tile59 As Long
Dim Tile60 As Long
Dim Tile61 As Long
Dim Tile62 As Long
Dim Tile63 As Long
Dim Tile64 As Long
Dim Tile65 As Long
Dim Tile66 As Long
Dim Tile67 As Long
Dim Tile68 As Long
Dim Tile69 As Long
Dim Tile70 As Long
Dim Tile71 As Long
Dim Tile72 As Long
Dim Tile73 As Long
Dim Tile74 As Long
Dim Tile75 As Long
Dim Tile76 As Long
Dim Tile77 As Long
Dim Tile78 As Long
Dim Tile79 As Long
Dim Tile80 As Long
Dim Tile81 As Long
Dim Tile82 As Long
Dim Tile83 As Long
Dim Tile84 As Long
Dim Tile85 As Long
Dim Tile86 As Long
Dim Tile87 As Long
Dim Tile88 As Long
Dim Tile89 As Long
Dim Tile90 As Long
Dim Tile91 As Long
Dim Tile92 As Long
Dim Tile93 As Long
Dim Tile94 As Long
Dim Tile95 As Long
Dim Tile96 As Long
Dim Tile97 As Long
Dim Tile98 As Long
Dim Tile99 As Long
Dim Tile100 As Long
Dim Tile101 As Long
Dim Tile102 As Long
Dim Tile103 As Long
Dim Tile104 As Long
Dim Tile105 As Long
Dim Tile106 As Long
Dim Tile107 As Long
Dim Tile108 As Long
Dim Tile109 As Long
Dim Tile110 As Long
Dim Tile111 As Long
Dim Tile112 As Long
Dim Tile113 As Long
Dim Tile114 As Long
Dim Tile115 As Long
Dim Tile116 As Long
Dim Tile117 As Long
Dim Tile118 As Long
Dim Tile119 As Long
Dim Tile120 As Long
Dim Tile121 As Long
Dim Tile122 As Long
Dim Tile123 As Long
Dim Tile124 As Long
Dim Tile125 As Long
Dim Tile126 As Long
Dim Tile127 As Long
Dim Tile128 As Long
Dim Tile129 As Long
Dim Tile130 As Long
Dim Tile131 As Long
Dim Tile132 As Long
Dim Tile133 As Long
Dim Tile134 As Long
Dim Tile135 As Long
Dim Tile136 As Long
Dim Tile137 As Long
Dim Tile138 As Long
Dim Tile139 As Long
Dim Tile140 As Long
Dim Tile141 As Long
Dim Tile142 As Long
Dim Tile143 As Long
Dim xScale1(100) As Integer
Dim yScale1(100) As Integer
Dim Pix As Integer
Dim j As Integer
Dim k As Integer
Dim p As Integer
Dim xDrawD(100) As Double
Dim yDrawD(100) As Double
Dim xScale2 As Integer
Dim yScale2 As Integer

'Adding the masks together and loading
If ROINum = 1 Then
    'Load the mask
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\ROIMask1.tif", BigMask
ElseIf ROINum = 2 Then
    'add masks together
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\Journals to Add ROIs\Two ROIs\add2rois.jnl"
    'Load the mask
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\TileTestMasked2.tif", BigMask
ElseIf ROINum = 3 Then
    'add masks together
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\Journals to Add ROIs\Three ROIs\add3rois.jnl"
    'Load the mask
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\TileTestMasked2.tif", BigMask
ElseIf ROINum = 4 Then
    'add masks together
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\Journals to Add ROIs\Four ROIs\add4rois.jnl"
    'Load the mask
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\TileTestMasked2.tif", BigMask
ElseIf ROINum = 5 Then
    'add masks together
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\Journals to Add ROIs\Five ROIs\add5rois.jnl"
    'Load the mask
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\TileTestMasked2.tif", BigMask
End If


'Initializing Global Bus to tell me which tiles to fire
TilesToFire(0) = 0
TilesToFire(1) = 0
TilesToFire(2) = 0
TilesToFire(3) = 0
TilesToFire(4) = 0
TilesToFire(5) = 0
TilesToFire(6) = 0
TilesToFire(7) = 0
TilesToFire(8) = 0
TilesToFire(9) = 0
TilesToFire(10) = 0
TilesToFire(11) = 0
TilesToFire(12) = 0
TilesToFire(13) = 0
TilesToFire(14) = 0
TilesToFire(15) = 0
TilesToFire(16) = 0
TilesToFire(17) = 0
TilesToFire(18) = 0
TilesToFire(19) = 0
TilesToFire(20) = 0
TilesToFire(21) = 0
TilesToFire(22) = 0
TilesToFire(23) = 0
TilesToFire(24) = 0
TilesToFire(25) = 0
TilesToFire(26) = 0
TilesToFire(27) = 0
TilesToFire(28) = 0
TilesToFire(29) = 0
TilesToFire(30) = 0
TilesToFire(31) = 0
TilesToFire(32) = 0
TilesToFire(33) = 0
TilesToFire(34) = 0
TilesToFire(35) = 0
TilesToFire(36) = 0
TilesToFire(37) = 0
TilesToFire(38) = 0
TilesToFire(39) = 0
TilesToFire(40) = 0
TilesToFire(41) = 0
TilesToFire(42) = 0
TilesToFire(43) = 0
TilesToFire(44) = 0
TilesToFire(45) = 0
TilesToFire(46) = 0
TilesToFire(47) = 0
TilesToFire(48) = 0
TilesToFire(49) = 0
TilesToFire(50) = 0
TilesToFire(51) = 0
TilesToFire(52) = 0
TilesToFire(53) = 0
TilesToFire(54) = 0
TilesToFire(55) = 0
TilesToFire(56) = 0
TilesToFire(57) = 0
TilesToFire(58) = 0
TilesToFire(59) = 0
TilesToFire(60) = 0
TilesToFire(61) = 0
TilesToFire(62) = 0
TilesToFire(63) = 0
TilesToFire(64) = 0
TilesToFire(65) = 0
TilesToFire(66) = 0
TilesToFire(67) = 0
TilesToFire(68) = 0
TilesToFire(69) = 0
TilesToFire(70) = 0
TilesToFire(71) = 0
TilesToFire(72) = 0
TilesToFire(73) = 0
TilesToFire(74) = 0
TilesToFire(75) = 0
TilesToFire(76) = 0
TilesToFire(77) = 0
TilesToFire(78) = 0
TilesToFire(79) = 0
TilesToFire(80) = 0
TilesToFire(81) = 0
TilesToFire(82) = 0
TilesToFire(83) = 0
TilesToFire(84) = 0
TilesToFire(85) = 0
TilesToFire(86) = 0
TilesToFire(87) = 0
TilesToFire(88) = 0
TilesToFire(89) = 0
TilesToFire(90) = 0
TilesToFire(91) = 0
TilesToFire(92) = 0
TilesToFire(93) = 0
TilesToFire(94) = 0
TilesToFire(95) = 0
TilesToFire(96) = 0
TilesToFire(97) = 0
TilesToFire(98) = 0
TilesToFire(99) = 0
TilesToFire(100) = 0
TilesToFire(101) = 0
TilesToFire(102) = 0
TilesToFire(103) = 0
TilesToFire(104) = 0
TilesToFire(105) = 0
TilesToFire(106) = 0
TilesToFire(107) = 0
TilesToFire(108) = 0
TilesToFire(109) = 0
TilesToFire(110) = 0
TilesToFire(111) = 0
TilesToFire(112) = 0
TilesToFire(113) = 0
TilesToFire(114) = 0
TilesToFire(115) = 0
TilesToFire(116) = 0
TilesToFire(117) = 0
TilesToFire(118) = 0
TilesToFire(119) = 0
TilesToFire(120) = 0
TilesToFire(121) = 0
TilesToFire(122) = 0
TilesToFire(123) = 0
TilesToFire(124) = 0
TilesToFire(125) = 0
TilesToFire(126) = 0
TilesToFire(127) = 0
TilesToFire(128) = 0
TilesToFire(129) = 0
TilesToFire(130) = 0
TilesToFire(131) = 0
TilesToFire(132) = 0
TilesToFire(133) = 0
TilesToFire(134) = 0
TilesToFire(135) = 0
TilesToFire(136) = 0
TilesToFire(137) = 0
TilesToFire(138) = 0
TilesToFire(139) = 0
TilesToFire(140) = 0
TilesToFire(141) = 0
TilesToFire(142) = 0
TilesToFire(143) = 0



'Some Scaling
For p = 0 To (MasterROICounter - 1)
    If Option3.Value = True Then ' 4x2 tiled image
        xDrawD(p) = CDbl(xDrawAllROIs(p) * 3.88)
        yDrawD(p) = CDbl(yDrawAllROIs(p) * 3.9)
    ElseIf Option4.Value = True Then '4x4 tiled image
        xDrawD(p) = CDbl(xDrawAllROIs(p) * 5.25)
        yDrawD(p) = CDbl(yDrawAllROIs(p) * 5.25)
    ElseIf Option5.Value = True Then '6x6 tiled image
        xDrawD(p) = CDbl(xDrawAllROIs(p) * 7.88)
        yDrawD(p) = CDbl(yDrawAllROIs(p) * 7.88)
    ElseIf Option6.Value = True Then '8x8 tiled image
        xDrawD(p) = CDbl(xDrawAllROIs(p) * 10.5)
        yDrawD(p) = CDbl(yDrawAllROIs(p) * 10.5)
    ElseIf Option7.Value = True Then '12x12 tiled image
        xDrawD(p) = CDbl(xDrawAllROIs(p) * 15.75)
        yDrawD(p) = CDbl(yDrawAllROIs(p) * 15.75)
    ElseIf Option10.Value = True Then '3x3 tiled image
        xDrawD(p) = CDbl(xDrawAllROIs(p) * 3.94)
        yDrawD(p) = CDbl(yDrawAllROIs(p) * 3.94)
    End If
Next p

'some debugging
Text4.Text = CStr(xDrawD(0))
Text6.Text = CStr(888)


'figuring out which tiles need to be removed from the big image
If Option3.Value = True Then ' 4x2 tiled image
    For a = 0 To (MasterROICounter - 1)
        If xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(0) = 1
        ElseIf xDrawD(a) > 512 And xDrawD(a) < 1025 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
             TilesToFire(1) = 1
        ElseIf xDrawD(a) > 1024 And xDrawD(a) < 1537 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(2) = 1
        ElseIf xDrawD(a) > 1536 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(3) = 1
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 0 And yDrawD(a) > 512 Then
            TilesToFire(4) = 1
        ElseIf xDrawD(a) > 512 And xDrawD(a) < 1025 And yDrawD(a) > 0 And yDrawD(a) > 512 Then
            TilesToFire(5) = 1
        ElseIf xDrawD(a) > 1024 And xDrawD(a) < 1537 And yDrawD(a) > 0 And yDrawD(a) > 512 Then
            TilesToFire(6) = 1
        ElseIf xDrawD(a) > 1536 And yDrawD(a) > 0 And yDrawD(a) > 512 Then
            TilesToFire(7) = 1
        End If
    Next a
End If


If Option10.Value = True Then '3x3 tiled image
    For a = 0 To (MasterROICounter - 1)
        If xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(0) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
             TilesToFire(1) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(2) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(3) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(4) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(5) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(6) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(7) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(8) = 1
        End If
    Next a
End If


If Option4.Value = True Then '4x4 tiled image
    For a = 0 To (MasterROICounter - 1)
        If xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(0) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
             TilesToFire(1) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(2) = 1
        ElseIf xDrawD(a) > 1535 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(3) = 1
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(4) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(5) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(6) = 1
        ElseIf xDrawD(a) > 1535 And yDrawD(a) > 0 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(7) = 1
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(8) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(9) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(10) = 1
        ElseIf xDrawD(a) > 1535 And yDrawD(a) > 0 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(11) = 1
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 0 And yDrawD(a) > 1535 Then
            TilesToFire(12) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 0 And yDrawD(a) > 1535 Then
            TilesToFire(13) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 0 And yDrawD(a) > 1535 Then
            TilesToFire(14) = 1
        ElseIf xDrawD(a) > 1535 And yDrawD(a) > 0 And yDrawD(a) > 1535 Then
            TilesToFire(15) = 1
        End If
    Next a
End If

If Option5.Value = True Then '6x6 tiled image
    For a = 0 To (MasterROICounter - 1)
        If xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(0) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
             TilesToFire(1) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(2) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(3) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(4) = 1
        ElseIf xDrawD(a) > 2599 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(5) = 1
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(6) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
             TilesToFire(7) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(8) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(9) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(10) = 1
        ElseIf xDrawD(a) > 2599 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(11) = 1
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(12) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
             TilesToFire(13) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(14) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(15) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(16) = 1
        ElseIf xDrawD(a) > 2599 And yDrawD(a) > 0 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(17) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(18) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
             TilesToFire(19) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(20) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(21) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(22) = 1
        ElseIf xDrawD(a) > 2599 And yDrawD(a) > 0 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(23) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(24) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
             TilesToFire(25) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(26) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(27) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(28) = 1
        ElseIf xDrawD(a) > 2599 And yDrawD(a) > 0 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(29) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 2559 Then
            TilesToFire(30) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 2559 Then
             TilesToFire(31) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 2559 Then
            TilesToFire(32) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 2559 Then
            TilesToFire(33) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 2559 Then
            TilesToFire(34) = 1
        ElseIf xDrawD(a) > 2599 And yDrawD(a) > 0 And yDrawD(a) > 2559 Then
            TilesToFire(35) = 1
        End If
    Next a
End If

If Option6.Value = True Then '8x8 tiled image
    For a = 0 To (MasterROICounter - 1)
        If xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(0) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(1) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(2) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(3) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(4) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(5) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(6) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(7) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(8) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(9) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(10) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(11) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(12) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(13) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(14) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(15) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(16) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(17) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(18) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(19) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(20) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(21) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(22) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(23) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(24) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(25) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(26) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(27) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(28) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(29) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(30) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(31) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(32) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(33) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(34) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(35) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(36) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(37) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(38) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(39) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(40) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(41) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(42) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(43) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(44) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(45) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(46) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(47) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(48) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(49) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(50) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(51) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(52) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(53) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(54) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(55) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(56) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(57) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(58) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(59) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(60) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(61) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(62) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(63) = 1
        End If
    Next a
End If

If Option7.Value = True Then '12x12 tiled image
    For a = 0 To (MasterROICounter - 1)
        If xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(0) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(1) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(2) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(3) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(4) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(5) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(6) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(7) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(8) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(9) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(10) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 0 And yDrawD(a) < 512 Then
            TilesToFire(11) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(12) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(13) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(14) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(15) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(16) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(17) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(18) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(19) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(20) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(21) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(22) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 511 And yDrawD(a) < 1024 Then
            TilesToFire(23) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(24) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(25) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(26) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(27) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(28) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(29) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(30) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(31) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(32) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(33) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(34) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 1023 And yDrawD(a) < 1536 Then
            TilesToFire(35) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(36) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(37) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(38) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(39) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(40) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(41) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(42) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(43) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(44) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(45) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(46) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 1535 And yDrawD(a) < 2048 Then
            TilesToFire(47) = 1
        
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(48) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(49) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(50) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(51) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(52) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(53) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(54) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(55) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(56) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(57) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(58) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 2047 And yDrawD(a) < 2560 Then
            TilesToFire(59) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(60) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(61) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(62) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(63) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(64) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(65) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(66) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(67) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(68) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(69) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(70) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 2559 And yDrawD(a) < 3072 Then
            TilesToFire(71) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(72) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(73) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(74) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(75) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(76) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(77) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(78) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(79) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(80) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(81) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(81) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 3071 And yDrawD(a) < 3584 Then
            TilesToFire(83) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(84) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(85) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(86) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(87) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(88) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(89) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(90) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(91) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(92) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(93) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(94) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 3583 And yDrawD(a) < 4096 Then
            TilesToFire(95) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(96) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(97) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(98) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(99) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(100) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(101) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(102) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(103) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(104) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(105) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(106) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 4095 And yDrawD(a) < 4608 Then
            TilesToFire(107) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(108) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(109) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(110) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(111) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(112) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(113) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(114) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(115) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(116) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(117) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(118) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 4607 And yDrawD(a) < 5120 Then
            TilesToFire(119) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(120) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(121) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(122) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(123) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(124) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(125) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(126) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(127) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(128) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(129) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(130) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 5119 And yDrawD(a) < 5632 Then
            TilesToFire(131) = 1
            
        ElseIf xDrawD(a) > 0 And xDrawD(a) < 512 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(132) = 1
        ElseIf xDrawD(a) > 511 And xDrawD(a) < 1024 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(133) = 1
        ElseIf xDrawD(a) > 1023 And xDrawD(a) < 1536 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(134) = 1
        ElseIf xDrawD(a) > 1535 And xDrawD(a) < 2048 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(135) = 1
        ElseIf xDrawD(a) > 2047 And xDrawD(a) < 2560 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(136) = 1
        ElseIf xDrawD(a) > 2599 And xDrawD(a) < 3072 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(137) = 1
        ElseIf xDrawD(a) > 3071 And xDrawD(a) < 3584 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(138) = 1
        ElseIf xDrawD(a) > 3583 And xDrawD(a) < 4096 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(139) = 1
        ElseIf xDrawD(a) > 4095 And xDrawD(a) < 4608 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(140) = 1
        ElseIf xDrawD(a) > 4607 And xDrawD(a) < 5120 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(141) = 1
        ElseIf xDrawD(a) > 5119 And xDrawD(a) < 5632 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(142) = 1
        ElseIf xDrawD(a) > 5631 And yDrawD(a) > 5631 And yDrawD(a) < 6144 Then
            TilesToFire(143) = 1
            
        End If
    Next a
End If



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'pulling out the individual tiles and saving them to the directory of the big mask

'from the mosaic

If Option3.Value = True Then ' 4x2 tiled image
    
    xScale1(0) = 0
    yScale1(0) = 0
    xScale1(1) = 512
    yScale1(1) = 0
    xScale1(2) = 1024
    yScale1(2) = 0
    xScale1(3) = 1536
    yScale1(3) = 0
    xScale1(4) = 0
    yScale1(4) = 512
    xScale1(5) = 512
    yScale1(5) = 512
    xScale1(6) = 1024
    yScale1(6) = 512
    xScale1(7) = 1536
    yScale1(7) = 512

    If TilesToFire(0) = 1 Then
        'Create image
        PubMM.CreateImage 512, 512, 16, "Tile0", Tile0
    
        'Creating image
        For j = 0 To 511
        For k = 0 To 511
            PubMM.ReadPixel BigMask, j + xScale1(0), k + yScale1(0), Pix
            PubMM.WritePixel Tile0, j, k, Pix
        Next k
        Next j
    
        'saving
        PubMM.SaveImage Tile0, "D:\Users\JohnE\VisualBasicTests\Tile0.tif", False, 3
    
        'Close image
        PubMM.CloseImage Tile0
    End If
    
    If TilesToFire(1) = 1 Then
        'Create image
        PubMM.CreateImage 512, 512, 16, "Tile1", Tile1
    
        'Creating image
        For j = 0 To 511
        For k = 0 To 511
            PubMM.ReadPixel BigMask, j + xScale1(1), k + yScale1(1), Pix
            PubMM.WritePixel Tile1, j, k, Pix
        Next k
        Next j

        'saving
        PubMM.SaveImage Tile1, "D:\Users\JohnE\VisualBasicTests\Tile1.tif", False, 3
    
        'Close image
        PubMM.CloseImage Tile1
    End If
    
    If TilesToFire(2) = 1 Then
        'Create image
        PubMM.CreateImage 512, 512, 16, "Tile2", Tile2
    
        'Creating image
        For j = 0 To 511
        For k = 0 To 511
            PubMM.ReadPixel BigMask, j + xScale1(2), k + yScale1(2), Pix
            PubMM.WritePixel Tile2, j, k, Pix
        Next k
        Next j
    
        'saving
        PubMM.SaveImage Tile2, "D:\Users\JohnE\VisualBasicTests\Tile2.tif", False, 3
    
        'Close image
        PubMM.CloseImage Tile2
    End If

    If TilesToFire(3) = 1 Then
        'Create image
        PubMM.CreateImage 512, 512, 16, "Tile3", Tile3
        
        'Creating image
        For j = 0 To 511
        For k = 0 To 511
           PubMM.ReadPixel BigMask, j + xScale1(3), k + yScale1(3), Pix
            PubMM.WritePixel Tile3, j, k, Pix
        Next k
        Next j
    
        'saving
        PubMM.SaveImage Tile3, "D:\Users\JohnE\VisualBasicTests\Tile3.tif", False, 3
    
        'Close image
        PubMM.CloseImage Tile3
    End If

    If TilesToFire(4) = 1 Then
        'Create image
        PubMM.CreateImage 512, 512, 16, "Tile4", Tile4
    
        'Creating image
        For j = 0 To 511
        For k = 0 To 511
            PubMM.ReadPixel BigMask, j + xScale1(4), k + yScale1(4), Pix
            PubMM.WritePixel Tile4, j, k, Pix
        Next k
        Next j
    
        'saving
        PubMM.SaveImage Tile4, "D:\Users\JohnE\VisualBasicTests\Tile4.tif", False, 3
    
        'Close image
        PubMM.CloseImage Tile4
    End If

    If TilesToFire(5) = 1 Then
        'Create image
        PubMM.CreateImage 512, 512, 16, "Tile5", Tile5

        'Creating image
        For j = 0 To 511
        For k = 0 To 511
            PubMM.ReadPixel BigMask, j + xScale1(5), k + yScale1(5), Pix
            PubMM.WritePixel Tile5, j, k, Pix
        Next k
        Next j

        'saving
        PubMM.SaveImage Tile5, "D:\Users\JohnE\VisualBasicTests\Tile5.tif", False, 3
    
        'Close image
         PubMM.CloseImage Tile5
    End If

    If TilesToFire(6) = 1 Then
        'Create image
        PubMM.CreateImage 512, 512, 16, "Tile6", Tile6
    
        'Creating image
        For j = 0 To 511
        For k = 0 To 511
            PubMM.ReadPixel BigMask, j + xScale1(6), k + yScale1(6), Pix
            PubMM.WritePixel Tile6, j, k, Pix
        Next k
        Next j
    
        'saving
        PubMM.SaveImage Tile6, "D:\Users\JohnE\VisualBasicTests\Tile6.tif", False, 3
    
        'Close image
        PubMM.CloseImage Tile6
    End If

    If TilesToFire(7) = 1 Then
        'Create image
        PubMM.CreateImage 512, 512, 16, "Tile7", Tile7
    
        'Creating image
        For j = 0 To 511
        For k = 0 To 511
            PubMM.ReadPixel BigMask, j + xScale1(7), k + yScale1(7), Pix
            PubMM.WritePixel Tile7, j, k, Pix
        Next k
        Next j
        
        'saving
        PubMM.SaveImage Tile7, "D:\Users\JohnE\VisualBasicTests\Tile7.tif", False, 3

        'Close image
        PubMM.CloseImage Tile7
    End If

    'Close the big image
    PubMM.CloseImage BigMask
End If

Dim Countx As Integer

'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
'''''''''3x3 tiled image

If Option10.Value = True Then

    'Some Counter
    Countx = 0
    
    For i = 0 To 8

        'xScale
        xScale = Countx * 512
            
        'x counter
        Countx = Countx + 1
        If Countx = 3 Then
            Countx = 0
        End If

        'yScale
        If i < 3 Then
            yScale = 0
        ElseIf i > 2 And i < 6 Then
            yScale = 512
        ElseIf i > 5 And i < 9 Then
            yScale = 1024
        End If
        
        'Making the tile
        If i = 0 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile0", Tile0
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile0, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile0, "D:\Users\JohnE\VisualBasicTests\Tile0.tif", False, 3

            'Close image
            PubMM.CloseImage Tile0
            
        ElseIf i = 1 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile1", Tile1
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile1, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile1, "D:\Users\JohnE\VisualBasicTests\Tile1.tif", False, 3

            'Close image
            PubMM.CloseImage Tile1
            
        ElseIf i = 2 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile2", Tile2
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile2, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile2, "D:\Users\JohnE\VisualBasicTests\Tile2.tif", False, 3

            'Close image
            PubMM.CloseImage Tile2
            
        ElseIf i = 3 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile3", Tile3
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile3, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile3, "D:\Users\JohnE\VisualBasicTests\Tile3.tif", False, 3

            'Close image
            PubMM.CloseImage Tile3
            
        ElseIf i = 4 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile4", Tile4
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile4, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile4, "D:\Users\JohnE\VisualBasicTests\Tile4.tif", False, 3

            'Close image
            PubMM.CloseImage Tile4
            
        ElseIf i = 5 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile5", Tile5
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile5, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile5, "D:\Users\JohnE\VisualBasicTests\Tile5.tif", False, 3

            'Close image
            PubMM.CloseImage Tile5
            
        ElseIf i = 6 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile6", Tile6
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile6, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile6, "D:\Users\JohnE\VisualBasicTests\Tile6.tif", False, 3

            'Close image
            PubMM.CloseImage Tile6
            
        ElseIf i = 7 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile7", Tile7
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile7, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile7, "D:\Users\JohnE\VisualBasicTests\Tile7.tif", False, 3

            'Close image
            PubMM.CloseImage Tile7
            
        ElseIf i = 8 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile8", Tile8
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile8, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile8, "D:\Users\JohnE\VisualBasicTests\Tile8.tif", False, 3

            'Close image
            PubMM.CloseImage Tile8


    End If

    Next i
    
    'Close image
    PubMM.CloseImage BigMask
        
End If

'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
'''''''''4x4 tiled image

If Option4.Value = True Then

    'Some Counter
    Countx = 0
    
    For i = 0 To 15

        'xScale
        xScale = Countx * 512
            
        'x counter
        Countx = Countx + 1
        If Countx = 4 Then
            Countx = 0
        End If

        'yScale
        If i < 4 Then
            yScale = 0
        ElseIf i > 3 And i < 8 Then
            yScale = 512
        ElseIf i > 7 And i < 12 Then
            yScale = 1024
        ElseIf i > 11 Then
            yScale = 1536
        End If
        
        'Making the tile
        If i = 0 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile0", Tile0
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile0, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile0, "D:\Users\JohnE\VisualBasicTests\Tile0.tif", False, 3

            'Close image
            PubMM.CloseImage Tile0
            
        ElseIf i = 1 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile1", Tile1
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile1, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile1, "D:\Users\JohnE\VisualBasicTests\Tile1.tif", False, 3

            'Close image
            PubMM.CloseImage Tile1
            
        ElseIf i = 2 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile2", Tile2
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile2, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile2, "D:\Users\JohnE\VisualBasicTests\Tile2.tif", False, 3

            'Close image
            PubMM.CloseImage Tile2
            
        ElseIf i = 3 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile3", Tile3
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile3, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile3, "D:\Users\JohnE\VisualBasicTests\Tile3.tif", False, 3

            'Close image
            PubMM.CloseImage Tile3
            
        ElseIf i = 4 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile4", Tile4
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile4, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile4, "D:\Users\JohnE\VisualBasicTests\Tile4.tif", False, 3

            'Close image
            PubMM.CloseImage Tile4
            
        ElseIf i = 5 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile5", Tile5
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile5, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile5, "D:\Users\JohnE\VisualBasicTests\Tile5.tif", False, 3

            'Close image
            PubMM.CloseImage Tile5
            
        ElseIf i = 6 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile6", Tile6
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile6, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile6, "D:\Users\JohnE\VisualBasicTests\Tile6.tif", False, 3

            'Close image
            PubMM.CloseImage Tile6
            
        ElseIf i = 7 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile7", Tile7
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile7, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile7, "D:\Users\JohnE\VisualBasicTests\Tile7.tif", False, 3

            'Close image
            PubMM.CloseImage Tile7
            
        ElseIf i = 8 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile8", Tile8
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile8, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile8, "D:\Users\JohnE\VisualBasicTests\Tile8.tif", False, 3

            'Close image
            PubMM.CloseImage Tile8
            
        ElseIf i = 9 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile9", Tile9
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile9, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile9, "D:\Users\JohnE\VisualBasicTests\Tile9.tif", False, 3

            'Close image
            PubMM.CloseImage Tile9
            
        ElseIf i = 10 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile10", Tile10
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile10, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile10, "D:\Users\JohnE\VisualBasicTests\Tile10.tif", False, 3

            'Close image
            PubMM.CloseImage Tile10
            
        ElseIf i = 11 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile11", Tile11
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile11, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile11, "D:\Users\JohnE\VisualBasicTests\Tile11.tif", False, 3

            'Close image
            PubMM.CloseImage Tile11
            
        ElseIf i = 12 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile12", Tile12
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile12, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile12, "D:\Users\JohnE\VisualBasicTests\Tile12.tif", False, 3

            'Close image
            PubMM.CloseImage Tile12
            
        ElseIf i = 13 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile13", Tile13
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile13, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile13, "D:\Users\JohnE\VisualBasicTests\Tile13.tif", False, 3

            'Close image
            PubMM.CloseImage Tile13
            
        ElseIf i = 14 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile14", Tile14
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile14, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile14, "D:\Users\JohnE\VisualBasicTests\Tile14.tif", False, 3

            'Close image
            PubMM.CloseImage Tile14
            
        ElseIf i = 15 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile15", Tile15
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale, k + yScale, Pix
                PubMM.WritePixel Tile15, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile15, "D:\Users\JohnE\VisualBasicTests\Tile15.tif", False, 3

            'Close image
            PubMM.CloseImage Tile15
            
        End If

    Next i
    
    'Close image
    PubMM.CloseImage BigMask
        
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

If Option5.Value = True Then '6x6 tiled image

 'intialize counter
    Countx = 0
    xScale2 = 0
    yScale2 = 0
     
    '6x6 mosaic
    For i = 0 To 35
    
            'xScale
            xScale2 = Countx * 512
            
            'x counter
            Countx = Countx + 1
            If Countx = 6 Then
                Countx = 0
            End If
            
            'yScale
            If i < 6 Then
                yScale2 = 0
            ElseIf i > 5 And i < 12 Then
                yScale2 = 512
            ElseIf i > 11 And i < 18 Then
                yScale2 = 1024
            ElseIf i > 17 And i < 24 Then
                yScale2 = 1536
            ElseIf i > 23 And i < 30 Then
                yScale2 = 2048
            ElseIf i > 29 And i < 36 Then
                yScale2 = 2560
            End If
                
        If i = 0 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile0", Tile0
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile0, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile0, "D:\Users\JohnE\VisualBasicTests\Tile0.tif", False, 3

            'Close image
            PubMM.CloseImage Tile0
            
        ElseIf i = 1 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile1", Tile1
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile1, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile1, "D:\Users\JohnE\VisualBasicTests\Tile1.tif", False, 3

            'Close image
            PubMM.CloseImage Tile1
            
        ElseIf i = 2 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile2", Tile2
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile2, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile2, "D:\Users\JohnE\VisualBasicTests\Tile2.tif", False, 3

            'Close image
            PubMM.CloseImage Tile2
            
        ElseIf i = 3 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile3", Tile3
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile3, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile3, "D:\Users\JohnE\VisualBasicTests\Tile3.tif", False, 3

            'Close image
            PubMM.CloseImage Tile3
            
        ElseIf i = 4 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile4", Tile4
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile4, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile4, "D:\Users\JohnE\VisualBasicTests\Tile4.tif", False, 3

            'Close image
            PubMM.CloseImage Tile4
            
        ElseIf i = 5 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile5", Tile5
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile5, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile5, "D:\Users\JohnE\VisualBasicTests\Tile5.tif", False, 3

            'Close image
            PubMM.CloseImage Tile5
            
        ElseIf i = 6 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile6", Tile6
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile6, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile6, "D:\Users\JohnE\VisualBasicTests\Tile6.tif", False, 3

            'Close image
            PubMM.CloseImage Tile6
            
        ElseIf i = 7 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile7", Tile7
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile7, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile7, "D:\Users\JohnE\VisualBasicTests\Tile7.tif", False, 3

            'Close image
            PubMM.CloseImage Tile7
            
        ElseIf i = 8 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile8", Tile8
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile8, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile8, "D:\Users\JohnE\VisualBasicTests\Tile8.tif", False, 3

            'Close image
            PubMM.CloseImage Tile8
            
        ElseIf i = 9 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile9", Tile9
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile9, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile9, "D:\Users\JohnE\VisualBasicTests\Tile9.tif", False, 3

            'Close image
            PubMM.CloseImage Tile9
            
        ElseIf i = 10 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile10", Tile10
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile10, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile10, "D:\Users\JohnE\VisualBasicTests\Tile10.tif", False, 3

            'Close image
            PubMM.CloseImage Tile10
            
        ElseIf i = 11 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile11", Tile11
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile11, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile11, "D:\Users\JohnE\VisualBasicTests\Tile11.tif", False, 3

            'Close image
            PubMM.CloseImage Tile11
            
        ElseIf i = 12 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile12", Tile12
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile12, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile12, "D:\Users\JohnE\VisualBasicTests\Tile12.tif", False, 3

            'Close image
            PubMM.CloseImage Tile12
            
        ElseIf i = 13 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile13", Tile13
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile13, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile13, "D:\Users\JohnE\VisualBasicTests\Tile13.tif", False, 3

            'Close image
            PubMM.CloseImage Tile13
            
        ElseIf i = 14 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile14", Tile14
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile14, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile14, "D:\Users\JohnE\VisualBasicTests\Tile14.tif", False, 3

            'Close image
            PubMM.CloseImage Tile14
            
        ElseIf i = 15 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile15", Tile15
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile15, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile15, "D:\Users\JohnE\VisualBasicTests\Tile15.tif", False, 3
            
            'Close image
            PubMM.CloseImage Tile15
            
        ElseIf i = 16 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile16", Tile16
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile16, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile16, "D:\Users\JohnE\VisualBasicTests\Tile16.tif", False, 3

            'Close image
            PubMM.CloseImage Tile16
            
        ElseIf i = 17 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile17", Tile17
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile17, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile17, "D:\Users\JohnE\VisualBasicTests\Tile17.tif", False, 3

            'Close image
            PubMM.CloseImage Tile17
            
        ElseIf i = 18 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile18", Tile18
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile2, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile18, "D:\Users\JohnE\VisualBasicTests\Tile18.tif", False, 3

            'Close image
            PubMM.CloseImage Tile18
            
        ElseIf i = 19 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile19", Tile19
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile19, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile19, "D:\Users\JohnE\VisualBasicTests\Tile19.tif", False, 3

            'Close image
            PubMM.CloseImage Tile19
            
        ElseIf i = 20 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile20", Tile20
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile20, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile20, "D:\Users\JohnE\VisualBasicTests\Tile20.tif", False, 3

            'Close image
            PubMM.CloseImage Tile20
            
        ElseIf i = 21 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile21", Tile21
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile21, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile21, "D:\Users\JohnE\VisualBasicTests\Tile21.tif", False, 3

            'Close image
            PubMM.CloseImage Tile21
            
        ElseIf i = 22 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile22", Tile22
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile22, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile22, "D:\Users\JohnE\VisualBasicTests\Tile22.tif", False, 3

            'Close image
            PubMM.CloseImage Tile22
            
        ElseIf i = 23 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile23", Tile23
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile23, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile23, "D:\Users\JohnE\VisualBasicTests\Tile23.tif", False, 3

            'Close image
            PubMM.CloseImage Tile23
            
        ElseIf i = 24 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile24", Tile24
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile24, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile24, "D:\Users\JohnE\VisualBasicTests\Tile24.tif", False, 3

            'Close image
            PubMM.CloseImage Tile24
            
        ElseIf i = 25 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile25", Tile25
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile25, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile25, "D:\Users\JohnE\VisualBasicTests\Tile25.tif", False, 3

            'Close image
            PubMM.CloseImage Tile25
            
        ElseIf i = 26 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile26", Tile26
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile26, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile26, "D:\Users\JohnE\VisualBasicTests\Tile26.tif", False, 3

            'Close image
            PubMM.CloseImage Tile26
            
        ElseIf i = 27 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile27", Tile27
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile27, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile27, "D:\Users\JohnE\VisualBasicTests\Tile27.tif", False, 3

            'Close image
            PubMM.CloseImage Tile27
            
        ElseIf i = 28 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile28", Tile28
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile28, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile28, "D:\Users\JohnE\VisualBasicTests\Tile28.tif", False, 3

            'Close image
            PubMM.CloseImage Tile28
            
        ElseIf i = 29 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile29", Tile29
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile29, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile29, "D:\Users\JohnE\VisualBasicTests\Tile29.tif", False, 3

            'Close image
            PubMM.CloseImage Tile29
            
        ElseIf i = 30 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile30", Tile30
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile30, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile30, "D:\Users\JohnE\VisualBasicTests\Tile30.tif", False, 3

            'Close image
            PubMM.CloseImage Tile30
            
        ElseIf i = 31 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile31", Tile31
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile31, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile31, "D:\Users\JohnE\VisualBasicTests\Tile31.tif", False, 3

            'Close image
            PubMM.CloseImage Tile31
            
        ElseIf i = 32 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile32", Tile32
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile32, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile32, "D:\Users\JohnE\VisualBasicTests\Tile32.tif", False, 3

            'Close image
            PubMM.CloseImage Tile32
            
        ElseIf i = 33 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile33", Tile33
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile33, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile33, "D:\Users\JohnE\VisualBasicTests\Tile33.tif", False, 3

            'Close image
            PubMM.CloseImage Tile33
            
        ElseIf i = 34 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile34", Tile34
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile34, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile34, "D:\Users\JohnE\VisualBasicTests\Tile34.tif", False, 3

            'Close image
            PubMM.CloseImage Tile34
            
        ElseIf i = 35 And TilesToFire(i) = 1 Then
        
            Text4.Text = CStr(xScale)
            Text6.Text = CStr(yScale)
        
            PubMM.CreateImage 512, 512, 16, "Tile35", Tile35
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile35, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile35, "D:\Users\JohnE\VisualBasicTests\Tile35.tif", False, 3

            'Close image
            PubMM.CloseImage Tile35
            
        End If
           
    Next i
    
    'Close the big image
    PubMM.CloseImage BigMask
    
End If

'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
''''''''''''8x8 tiled image

If Option6.Value = True Then

 'intialize counter
    Countx = 0
    xScale2 = 0
    yScale2 = 0
     
    '6x6 mosaic
    For i = 0 To 63
    
            'xScale
            xScale2 = Countx * 512
            
            'x counter
            Countx = Countx + 1
            If Countx = 8 Then
                Countx = 0
            End If
            
            'yScale
            If i < 8 Then
                yScale2 = 0
            ElseIf i > 7 And i < 16 Then
                yScale2 = 512
            ElseIf i > 15 And i < 24 Then
                yScale2 = 1024
            ElseIf i > 23 And i < 32 Then
                yScale2 = 1536
            ElseIf i > 31 And i < 40 Then
                yScale2 = 2048
            ElseIf i > 39 And i < 48 Then
                yScale2 = 2560
            ElseIf i > 47 And i < 56 Then
                yScale2 = 3072
            ElseIf i > 55 And i < 64 Then
                yScale2 = 3584
            End If
                
        If i = 0 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile0", Tile0
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile0, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile0, "D:\Users\JohnE\VisualBasicTests\Tile0.tif", False, 3

            'Close image
            PubMM.CloseImage Tile0
            
        ElseIf i = 1 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile1", Tile1
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile1, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile1, "D:\Users\JohnE\VisualBasicTests\Tile1.tif", False, 3

            'Close image
            PubMM.CloseImage Tile1
            
        ElseIf i = 2 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile2", Tile2
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile2, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile2, "D:\Users\JohnE\VisualBasicTests\Tile2.tif", False, 3

            'Close image
            PubMM.CloseImage Tile2
            
        ElseIf i = 3 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile3", Tile3
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile3, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile3, "D:\Users\JohnE\VisualBasicTests\Tile3.tif", False, 3

            'Close image
            PubMM.CloseImage Tile3
            
        ElseIf i = 4 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile4", Tile4
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile4, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile4, "D:\Users\JohnE\VisualBasicTests\Tile4.tif", False, 3

            'Close image
            PubMM.CloseImage Tile4
            
        ElseIf i = 5 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile5", Tile5
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile5, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile5, "D:\Users\JohnE\VisualBasicTests\Tile5.tif", False, 3

            'Close image
            PubMM.CloseImage Tile5
            
        ElseIf i = 6 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile6", Tile6
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile6, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile6, "D:\Users\JohnE\VisualBasicTests\Tile6.tif", False, 3

            'Close image
            PubMM.CloseImage Tile6
            
        ElseIf i = 7 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile7", Tile7
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile7, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile7, "D:\Users\JohnE\VisualBasicTests\Tile7.tif", False, 3

            'Close image
            PubMM.CloseImage Tile7
            
        ElseIf i = 8 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile8", Tile8
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile8, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile8, "D:\Users\JohnE\VisualBasicTests\Tile8.tif", False, 3

            'Close image
            PubMM.CloseImage Tile8
            
        ElseIf i = 9 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile9", Tile9
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile9, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile9, "D:\Users\JohnE\VisualBasicTests\Tile9.tif", False, 3

            'Close image
            PubMM.CloseImage Tile9
            
        ElseIf i = 10 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile10", Tile10
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile10, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile10, "D:\Users\JohnE\VisualBasicTests\Tile10.tif", False, 3

            'Close image
            PubMM.CloseImage Tile10
            
        ElseIf i = 11 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile11", Tile11
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile11, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile11, "D:\Users\JohnE\VisualBasicTests\Tile11.tif", False, 3

            'Close image
            PubMM.CloseImage Tile11
            
        ElseIf i = 12 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile12", Tile12
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile12, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile12, "D:\Users\JohnE\VisualBasicTests\Tile12.tif", False, 3

            'Close image
            PubMM.CloseImage Tile12
            
        ElseIf i = 13 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile13", Tile13
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile13, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile13, "D:\Users\JohnE\VisualBasicTests\Tile13.tif", False, 3

            'Close image
            PubMM.CloseImage Tile13
            
        ElseIf i = 14 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile14", Tile14
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile14, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile14, "D:\Users\JohnE\VisualBasicTests\Tile14.tif", False, 3

            'Close image
            PubMM.CloseImage Tile14
            
        ElseIf i = 15 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile15", Tile15
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile15, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile15, "D:\Users\JohnE\VisualBasicTests\Tile15.tif", False, 3
            
            'Close image
            PubMM.CloseImage Tile15
            
        ElseIf i = 16 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile16", Tile16
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile16, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile16, "D:\Users\JohnE\VisualBasicTests\Tile16.tif", False, 3

            'Close image
            PubMM.CloseImage Tile16
            
        ElseIf i = 17 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile17", Tile17
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile17, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile17, "D:\Users\JohnE\VisualBasicTests\Tile17.tif", False, 3

            'Close image
            PubMM.CloseImage Tile17
            
        ElseIf i = 18 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile18", Tile18
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile2, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile18, "D:\Users\JohnE\VisualBasicTests\Tile18.tif", False, 3

            'Close image
            PubMM.CloseImage Tile18
            
        ElseIf i = 19 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile19", Tile19
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile19, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile19, "D:\Users\JohnE\VisualBasicTests\Tile19.tif", False, 3

            'Close image
            PubMM.CloseImage Tile19
            
        ElseIf i = 20 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile20", Tile20
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile20, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile20, "D:\Users\JohnE\VisualBasicTests\Tile20.tif", False, 3

            'Close image
            PubMM.CloseImage Tile20
            
        ElseIf i = 21 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile21", Tile21
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile21, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile21, "D:\Users\JohnE\VisualBasicTests\Tile21.tif", False, 3

            'Close image
            PubMM.CloseImage Tile21
            
        ElseIf i = 22 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile22", Tile22
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile22, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile22, "D:\Users\JohnE\VisualBasicTests\Tile22.tif", False, 3

            'Close image
            PubMM.CloseImage Tile22
            
        ElseIf i = 23 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile23", Tile23
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile23, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile23, "D:\Users\JohnE\VisualBasicTests\Tile23.tif", False, 3

            'Close image
            PubMM.CloseImage Tile23
            
        ElseIf i = 24 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile24", Tile24
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile24, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile24, "D:\Users\JohnE\VisualBasicTests\Tile24.tif", False, 3

            'Close image
            PubMM.CloseImage Tile24
            
        ElseIf i = 25 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile25", Tile25
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile25, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile25, "D:\Users\JohnE\VisualBasicTests\Tile25.tif", False, 3

            'Close image
            PubMM.CloseImage Tile25
            
        ElseIf i = 26 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile26", Tile26
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile26, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile26, "D:\Users\JohnE\VisualBasicTests\Tile26.tif", False, 3

            'Close image
            PubMM.CloseImage Tile26
            
        ElseIf i = 27 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile27", Tile27
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile27, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile27, "D:\Users\JohnE\VisualBasicTests\Tile27.tif", False, 3

            'Close image
            PubMM.CloseImage Tile27
            
        ElseIf i = 28 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile28", Tile28
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile28, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile28, "D:\Users\JohnE\VisualBasicTests\Tile28.tif", False, 3

            'Close image
            PubMM.CloseImage Tile28
            
        ElseIf i = 29 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile29", Tile29
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile29, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile29, "D:\Users\JohnE\VisualBasicTests\Tile29.tif", False, 3

            'Close image
            PubMM.CloseImage Tile29
            
        ElseIf i = 30 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile30", Tile30
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile30, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile30, "D:\Users\JohnE\VisualBasicTests\Tile30.tif", False, 3

            'Close image
            PubMM.CloseImage Tile30
            
        ElseIf i = 31 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile31", Tile31
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile31, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile31, "D:\Users\JohnE\VisualBasicTests\Tile31.tif", False, 3

            'Close image
            PubMM.CloseImage Tile31
            
        ElseIf i = 32 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile32", Tile32
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile32, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile32, "D:\Users\JohnE\VisualBasicTests\Tile32.tif", False, 3

            'Close image
            PubMM.CloseImage Tile32
            
        ElseIf i = 33 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile33", Tile33
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile33, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile33, "D:\Users\JohnE\VisualBasicTests\Tile33.tif", False, 3

            'Close image
            PubMM.CloseImage Tile33
            
        ElseIf i = 34 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile34", Tile34
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile34, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile34, "D:\Users\JohnE\VisualBasicTests\Tile34.tif", False, 3

            'Close image
            PubMM.CloseImage Tile34
            
        ElseIf i = 35 And TilesToFire(i) = 1 Then
        
            Text4.Text = CStr(xScale)
            Text6.Text = CStr(yScale)
        
            PubMM.CreateImage 512, 512, 16, "Tile35", Tile35
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile35, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile35, "D:\Users\JohnE\VisualBasicTests\Tile35.tif", False, 3

            'Close image
            PubMM.CloseImage Tile35
            
         ElseIf i = 36 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile36", Tile36
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile36, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile36, "D:\Users\JohnE\VisualBasicTests\Tile36.tif", False, 3

            'Close image
            PubMM.CloseImage Tile36
            
        ElseIf i = 37 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile37", Tile37
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile37, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile37, "D:\Users\JohnE\VisualBasicTests\Tile37.tif", False, 3

            'Close image
            PubMM.CloseImage Tile37
            
        ElseIf i = 38 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile38", Tile38
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile38, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile38, "D:\Users\JohnE\VisualBasicTests\Tile38.tif", False, 3

            'Close image
            PubMM.CloseImage Tile38
            
        ElseIf i = 39 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile39", Tile39
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile39, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile39, "D:\Users\JohnE\VisualBasicTests\Tile39.tif", False, 3

            'Close image
            PubMM.CloseImage Tile39
            
        ElseIf i = 40 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile40", Tile40
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile40, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile40, "D:\Users\JohnE\VisualBasicTests\Tile40.tif", False, 3

            'Close image
            PubMM.CloseImage Tile40
            
        ElseIf i = 41 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile41", Tile41
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile41, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile41, "D:\Users\JohnE\VisualBasicTests\Tile41.tif", False, 3

            'Close image
            PubMM.CloseImage Tile41
            
        ElseIf i = 42 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile42", Tile42
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile42, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile42, "D:\Users\JohnE\VisualBasicTests\Tile42.tif", False, 3

            'Close image
            PubMM.CloseImage Tile42
            
        ElseIf i = 43 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile43", Tile43
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile43, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile43, "D:\Users\JohnE\VisualBasicTests\Tile43.tif", False, 3

            'Close image
            PubMM.CloseImage Tile43
            
        ElseIf i = 44 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile44", Tile44
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile44, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile44, "D:\Users\JohnE\VisualBasicTests\Tile44.tif", False, 3

            'Close image
            PubMM.CloseImage Tile44
            
        ElseIf i = 45 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile45", Tile45
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile45, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile45, "D:\Users\JohnE\VisualBasicTests\Tile45.tif", False, 3

            'Close image
            PubMM.CloseImage Tile45
            
        ElseIf i = 46 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile46", Tile46
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile46, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile46, "D:\Users\JohnE\VisualBasicTests\Tile46.tif", False, 3

            'Close image
            PubMM.CloseImage Tile46
            
        ElseIf i = 47 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile47", Tile47
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile47, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile47, "D:\Users\JohnE\VisualBasicTests\Tile47.tif", False, 3

            'Close image
            PubMM.CloseImage Tile47
            
        ElseIf i = 48 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile48", Tile48
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile48, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile48, "D:\Users\JohnE\VisualBasicTests\Tile48.tif", False, 3

            'Close image
            PubMM.CloseImage Tile48
            
        ElseIf i = 49 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile49", Tile49
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile49, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile49, "D:\Users\JohnE\VisualBasicTests\Tile49.tif", False, 3
            
            'Close image
            PubMM.CloseImage Tile49
            
        ElseIf i = 50 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile50", Tile50
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile50, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile50, "D:\Users\JohnE\VisualBasicTests\Tile50.tif", False, 3

            'Close image
            PubMM.CloseImage Tile50
            
        ElseIf i = 51 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile51", Tile51
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile51, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile51, "D:\Users\JohnE\VisualBasicTests\Tile51.tif", False, 3

            'Close image
            PubMM.CloseImage Tile51
            
        ElseIf i = 52 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile52", Tile52
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile52, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile52, "D:\Users\JohnE\VisualBasicTests\Tile52.tif", False, 3

            'Close image
            PubMM.CloseImage Tile52
            
        ElseIf i = 53 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile53", Tile53
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile53, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile53, "D:\Users\JohnE\VisualBasicTests\Tile53.tif", False, 3

            'Close image
            PubMM.CloseImage Tile53
            
        ElseIf i = 54 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile54", Tile54
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile54, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile54, "D:\Users\JohnE\VisualBasicTests\Tile54.tif", False, 3

            'Close image
            PubMM.CloseImage Tile54
            
        ElseIf i = 55 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile55", Tile55
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile55, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile55, "D:\Users\JohnE\VisualBasicTests\Tile55.tif", False, 3

            'Close image
            PubMM.CloseImage Tile55
            
        ElseIf i = 56 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile56", Tile56
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile56, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile56, "D:\Users\JohnE\VisualBasicTests\Tile56.tif", False, 3

            'Close image
            PubMM.CloseImage Tile56
            
        ElseIf i = 57 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile57", Tile57
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile57, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile57, "D:\Users\JohnE\VisualBasicTests\Tile57.tif", False, 3

            'Close image
            PubMM.CloseImage Tile57
            
        ElseIf i = 58 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile58", Tile58
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile58, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile58, "D:\Users\JohnE\VisualBasicTests\Tile58.tif", False, 3

            'Close image
            PubMM.CloseImage Tile58
            
        ElseIf i = 59 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile59", Tile59
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile59, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile59, "D:\Users\JohnE\VisualBasicTests\Tile59.tif", False, 3

            'Close image
            PubMM.CloseImage Tile59
            
        ElseIf i = 60 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile60", Tile60
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile60, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile60, "D:\Users\JohnE\VisualBasicTests\Tile60.tif", False, 3

            'Close image
            PubMM.CloseImage Tile60
            
        ElseIf i = 61 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile61", Tile61
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile61, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile61, "D:\Users\JohnE\VisualBasicTests\Tile61.tif", False, 3

            'Close image
            PubMM.CloseImage Tile61
            
        ElseIf i = 62 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile62", Tile62
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile62, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile62, "D:\Users\JohnE\VisualBasicTests\Tile62.tif", False, 3

            'Close image
            PubMM.CloseImage Tile62
            
        ElseIf i = 63 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile63", Tile63
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile63, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile63, "D:\Users\JohnE\VisualBasicTests\Tile63.tif", False, 3

            'Close image
            PubMM.CloseImage Tile63
            
        
        End If
        
       
    Next i


    'Close image
    PubMM.CloseImage BigMask


End If


''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
''''''''''''12x12 tiled image

If Option7.Value = True Then

 'intialize counter
    Countx = 0
    xScale2 = 0
    yScale2 = 0
     
    '6x6 mosaic
    For i = 0 To 143
    
            'xScale
            xScale2 = Countx * 512
            
            'x counter
            Countx = Countx + 1
            If Countx = 12 Then
                Countx = 0
            End If
            
            'yScale
            If i < 12 Then
                yScale2 = 0
            ElseIf i > 11 And i < 24 Then
                yScale2 = 512
            ElseIf i > 23 And i < 36 Then
                yScale2 = 1024
            ElseIf i > 35 And i < 48 Then
                yScale2 = 1536
            ElseIf i > 47 And i < 60 Then
                yScale2 = 2048
            ElseIf i > 59 And i < 72 Then
                yScale2 = 2560
            ElseIf i > 71 And i < 84 Then
                yScale2 = 3072
            ElseIf i > 83 And i < 96 Then
                yScale2 = 3584
            ElseIf i > 95 And i < 108 Then
                yScale2 = 4096
            ElseIf i > 107 And i < 120 Then
                yScale2 = 4608
            ElseIf i > 119 And i < 132 Then
                yScale2 = 5120
            ElseIf i > 131 And i < 144 Then
                yScale2 = 5632
            End If
                
        If i = 0 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile0", Tile0
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile0, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile0, "D:\Users\JohnE\VisualBasicTests\Tile0.tif", False, 3

            'Close image
            PubMM.CloseImage Tile0
            
        ElseIf i = 1 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile1", Tile1
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile1, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile1, "D:\Users\JohnE\VisualBasicTests\Tile1.tif", False, 3

            'Close image
            PubMM.CloseImage Tile1
            
        ElseIf i = 2 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile2", Tile2
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile2, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile2, "D:\Users\JohnE\VisualBasicTests\Tile2.tif", False, 3

            'Close image
            PubMM.CloseImage Tile2
            
        ElseIf i = 3 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile3", Tile3
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile3, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile3, "D:\Users\JohnE\VisualBasicTests\Tile3.tif", False, 3

            'Close image
            PubMM.CloseImage Tile3
            
        ElseIf i = 4 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile4", Tile4
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile4, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile4, "D:\Users\JohnE\VisualBasicTests\Tile4.tif", False, 3

            'Close image
            PubMM.CloseImage Tile4
            
        ElseIf i = 5 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile5", Tile5
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile5, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile5, "D:\Users\JohnE\VisualBasicTests\Tile5.tif", False, 3

            'Close image
            PubMM.CloseImage Tile5
            
        ElseIf i = 6 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile6", Tile6
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile6, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile6, "D:\Users\JohnE\VisualBasicTests\Tile6.tif", False, 3

            'Close image
            PubMM.CloseImage Tile6
            
        ElseIf i = 7 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile7", Tile7
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile7, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile7, "D:\Users\JohnE\VisualBasicTests\Tile7.tif", False, 3

            'Close image
            PubMM.CloseImage Tile7
            
        ElseIf i = 8 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile8", Tile8
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile8, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile8, "D:\Users\JohnE\VisualBasicTests\Tile8.tif", False, 3

            'Close image
            PubMM.CloseImage Tile8
            
        ElseIf i = 9 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile9", Tile9
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile9, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile9, "D:\Users\JohnE\VisualBasicTests\Tile9.tif", False, 3

            'Close image
            PubMM.CloseImage Tile9
            
        ElseIf i = 10 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile10", Tile10
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile10, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile10, "D:\Users\JohnE\VisualBasicTests\Tile10.tif", False, 3

            'Close image
            PubMM.CloseImage Tile10
            
        ElseIf i = 11 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile11", Tile11
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile11, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile11, "D:\Users\JohnE\VisualBasicTests\Tile11.tif", False, 3

            'Close image
            PubMM.CloseImage Tile11
            
        ElseIf i = 12 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile12", Tile12
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile12, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile12, "D:\Users\JohnE\VisualBasicTests\Tile12.tif", False, 3

            'Close image
            PubMM.CloseImage Tile12
            
        ElseIf i = 13 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile13", Tile13
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile13, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile13, "D:\Users\JohnE\VisualBasicTests\Tile13.tif", False, 3

            'Close image
            PubMM.CloseImage Tile13
            
        ElseIf i = 14 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile14", Tile14
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile14, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile14, "D:\Users\JohnE\VisualBasicTests\Tile14.tif", False, 3

            'Close image
            PubMM.CloseImage Tile14
            
        ElseIf i = 15 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile15", Tile15
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile15, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile15, "D:\Users\JohnE\VisualBasicTests\Tile15.tif", False, 3
            
            'Close image
            PubMM.CloseImage Tile15
            
        ElseIf i = 16 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile16", Tile16
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile16, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile16, "D:\Users\JohnE\VisualBasicTests\Tile16.tif", False, 3

            'Close image
            PubMM.CloseImage Tile16
            
        ElseIf i = 17 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile17", Tile17
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile17, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile17, "D:\Users\JohnE\VisualBasicTests\Tile17.tif", False, 3

            'Close image
            PubMM.CloseImage Tile17
            
        ElseIf i = 18 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile18", Tile18
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile2, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile18, "D:\Users\JohnE\VisualBasicTests\Tile18.tif", False, 3

            'Close image
            PubMM.CloseImage Tile18
            
        ElseIf i = 19 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile19", Tile19
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile19, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile19, "D:\Users\JohnE\VisualBasicTests\Tile19.tif", False, 3

            'Close image
            PubMM.CloseImage Tile19
            
        ElseIf i = 20 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile20", Tile20
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile20, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile20, "D:\Users\JohnE\VisualBasicTests\Tile20.tif", False, 3

            'Close image
            PubMM.CloseImage Tile20
            
        ElseIf i = 21 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile21", Tile21
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile21, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile21, "D:\Users\JohnE\VisualBasicTests\Tile21.tif", False, 3

            'Close image
            PubMM.CloseImage Tile21
            
        ElseIf i = 22 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile22", Tile22
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile22, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile22, "D:\Users\JohnE\VisualBasicTests\Tile22.tif", False, 3

            'Close image
            PubMM.CloseImage Tile22
            
        ElseIf i = 23 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile23", Tile23
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile23, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile23, "D:\Users\JohnE\VisualBasicTests\Tile23.tif", False, 3

            'Close image
            PubMM.CloseImage Tile23
            
        ElseIf i = 24 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile24", Tile24
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile24, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile24, "D:\Users\JohnE\VisualBasicTests\Tile24.tif", False, 3

            'Close image
            PubMM.CloseImage Tile24
            
        ElseIf i = 25 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile25", Tile25
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile25, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile25, "D:\Users\JohnE\VisualBasicTests\Tile25.tif", False, 3

            'Close image
            PubMM.CloseImage Tile25
            
        ElseIf i = 26 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile26", Tile26
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile26, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile26, "D:\Users\JohnE\VisualBasicTests\Tile26.tif", False, 3

            'Close image
            PubMM.CloseImage Tile26
            
        ElseIf i = 27 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile27", Tile27
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile27, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile27, "D:\Users\JohnE\VisualBasicTests\Tile27.tif", False, 3

            'Close image
            PubMM.CloseImage Tile27
            
        ElseIf i = 28 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile28", Tile28
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile28, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile28, "D:\Users\JohnE\VisualBasicTests\Tile28.tif", False, 3

            'Close image
            PubMM.CloseImage Tile28
            
        ElseIf i = 29 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile29", Tile29
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile29, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile29, "D:\Users\JohnE\VisualBasicTests\Tile29.tif", False, 3

            'Close image
            PubMM.CloseImage Tile29
            
        ElseIf i = 30 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile30", Tile30
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile30, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile30, "D:\Users\JohnE\VisualBasicTests\Tile30.tif", False, 3

            'Close image
            PubMM.CloseImage Tile30
            
        ElseIf i = 31 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile31", Tile31
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile31, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile31, "D:\Users\JohnE\VisualBasicTests\Tile31.tif", False, 3

            'Close image
            PubMM.CloseImage Tile31
            
        ElseIf i = 32 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile32", Tile32
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile32, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile32, "D:\Users\JohnE\VisualBasicTests\Tile32.tif", False, 3

            'Close image
            PubMM.CloseImage Tile32
            
        ElseIf i = 33 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile33", Tile33
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile33, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile33, "D:\Users\JohnE\VisualBasicTests\Tile33.tif", False, 3

            'Close image
            PubMM.CloseImage Tile33
            
        ElseIf i = 34 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile34", Tile34
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile34, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile34, "D:\Users\JohnE\VisualBasicTests\Tile34.tif", False, 3

            'Close image
            PubMM.CloseImage Tile34
            
        ElseIf i = 35 And TilesToFire(i) = 1 Then
        
            Text4.Text = CStr(xScale)
            Text6.Text = CStr(yScale)
        
            PubMM.CreateImage 512, 512, 16, "Tile35", Tile35
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile35, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile35, "D:\Users\JohnE\VisualBasicTests\Tile35.tif", False, 3

            'Close image
            PubMM.CloseImage Tile35
            
         ElseIf i = 36 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile36", Tile36
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile36, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile36, "D:\Users\JohnE\VisualBasicTests\Tile36.tif", False, 3

            'Close image
            PubMM.CloseImage Tile36
            
        ElseIf i = 37 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile37", Tile37
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile37, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile37, "D:\Users\JohnE\VisualBasicTests\Tile37.tif", False, 3

            'Close image
            PubMM.CloseImage Tile37
            
        ElseIf i = 38 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile38", Tile38
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile38, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile38, "D:\Users\JohnE\VisualBasicTests\Tile38.tif", False, 3

            'Close image
            PubMM.CloseImage Tile38
            
        ElseIf i = 39 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile39", Tile39
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile39, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile39, "D:\Users\JohnE\VisualBasicTests\Tile39.tif", False, 3

            'Close image
            PubMM.CloseImage Tile39
            
        ElseIf i = 40 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile40", Tile40
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile40, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile40, "D:\Users\JohnE\VisualBasicTests\Tile40.tif", False, 3

            'Close image
            PubMM.CloseImage Tile40
            
        ElseIf i = 41 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile41", Tile41
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile41, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile41, "D:\Users\JohnE\VisualBasicTests\Tile41.tif", False, 3

            'Close image
            PubMM.CloseImage Tile41
            
        ElseIf i = 42 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile42", Tile42
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile42, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile42, "D:\Users\JohnE\VisualBasicTests\Tile42.tif", False, 3

            'Close image
            PubMM.CloseImage Tile42
            
        ElseIf i = 43 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile43", Tile43
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile43, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile43, "D:\Users\JohnE\VisualBasicTests\Tile43.tif", False, 3

            'Close image
            PubMM.CloseImage Tile43
            
        ElseIf i = 44 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile44", Tile44
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile44, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile44, "D:\Users\JohnE\VisualBasicTests\Tile44.tif", False, 3

            'Close image
            PubMM.CloseImage Tile44
            
        ElseIf i = 45 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile45", Tile45
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile45, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile45, "D:\Users\JohnE\VisualBasicTests\Tile45.tif", False, 3

            'Close image
            PubMM.CloseImage Tile45
            
        ElseIf i = 46 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile46", Tile46
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile46, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile46, "D:\Users\JohnE\VisualBasicTests\Tile46.tif", False, 3

            'Close image
            PubMM.CloseImage Tile46
            
        ElseIf i = 47 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile47", Tile47
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile47, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile47, "D:\Users\JohnE\VisualBasicTests\Tile47.tif", False, 3

            'Close image
            PubMM.CloseImage Tile47
            
        ElseIf i = 48 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile48", Tile48
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile48, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile48, "D:\Users\JohnE\VisualBasicTests\Tile48.tif", False, 3

            'Close image
            PubMM.CloseImage Tile48
            
        ElseIf i = 49 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile49", Tile49
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile49, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile49, "D:\Users\JohnE\VisualBasicTests\Tile49.tif", False, 3
            
            'Close image
            PubMM.CloseImage Tile49
            
        ElseIf i = 50 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile50", Tile50
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile50, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile50, "D:\Users\JohnE\VisualBasicTests\Tile50.tif", False, 3

            'Close image
            PubMM.CloseImage Tile50
            
        ElseIf i = 51 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile51", Tile51
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile51, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile51, "D:\Users\JohnE\VisualBasicTests\Tile51.tif", False, 3

            'Close image
            PubMM.CloseImage Tile51
            
        ElseIf i = 52 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile52", Tile52
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile52, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile52, "D:\Users\JohnE\VisualBasicTests\Tile52.tif", False, 3

            'Close image
            PubMM.CloseImage Tile52
            
        ElseIf i = 53 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile53", Tile53
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile53, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile53, "D:\Users\JohnE\VisualBasicTests\Tile53.tif", False, 3

            'Close image
            PubMM.CloseImage Tile53
            
        ElseIf i = 54 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile54", Tile54
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile54, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile54, "D:\Users\JohnE\VisualBasicTests\Tile54.tif", False, 3

            'Close image
            PubMM.CloseImage Tile54
            
        ElseIf i = 55 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile55", Tile55
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile55, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile55, "D:\Users\JohnE\VisualBasicTests\Tile55.tif", False, 3

            'Close image
            PubMM.CloseImage Tile55
            
        ElseIf i = 56 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile56", Tile56
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile56, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile56, "D:\Users\JohnE\VisualBasicTests\Tile56.tif", False, 3

            'Close image
            PubMM.CloseImage Tile56
            
        ElseIf i = 57 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile57", Tile57
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile57, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile57, "D:\Users\JohnE\VisualBasicTests\Tile57.tif", False, 3

            'Close image
            PubMM.CloseImage Tile57
            
        ElseIf i = 58 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile58", Tile58
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile58, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile58, "D:\Users\JohnE\VisualBasicTests\Tile58.tif", False, 3

            'Close image
            PubMM.CloseImage Tile58
            
        ElseIf i = 59 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile59", Tile59
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile59, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile59, "D:\Users\JohnE\VisualBasicTests\Tile59.tif", False, 3

            'Close image
            PubMM.CloseImage Tile59
            
        ElseIf i = 60 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile60", Tile60
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile60, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile60, "D:\Users\JohnE\VisualBasicTests\Tile60.tif", False, 3

            'Close image
            PubMM.CloseImage Tile60
            
        ElseIf i = 61 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile61", Tile61
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile61, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile61, "D:\Users\JohnE\VisualBasicTests\Tile61.tif", False, 3

            'Close image
            PubMM.CloseImage Tile61
            
        ElseIf i = 62 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile62", Tile62
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile62, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile62, "D:\Users\JohnE\VisualBasicTests\Tile62.tif", False, 3

            'Close image
            PubMM.CloseImage Tile62
            
        ElseIf i = 63 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile63", Tile63
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile63, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile63, "D:\Users\JohnE\VisualBasicTests\Tile63.tif", False, 3

            'Close image
            PubMM.CloseImage Tile63
            
        ElseIf i = 64 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile64", Tile64
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile64, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile64, "D:\Users\JohnE\VisualBasicTests\Tile64.tif", False, 3

            'Close image
            PubMM.CloseImage Tile64

        ElseIf i = 65 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile65", Tile65
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile65, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile65, "D:\Users\JohnE\VisualBasicTests\Tile65.tif", False, 3

            'Close image
            PubMM.CloseImage Tile65
            
        ElseIf i = 66 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile66", Tile66
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile66, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile66, "D:\Users\JohnE\VisualBasicTests\Tile66.tif", False, 3

            'Close image
            PubMM.CloseImage Tile66

        ElseIf i = 67 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile67", Tile67
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile67, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile67, "D:\Users\JohnE\VisualBasicTests\Tile67.tif", False, 3

            'Close image
            PubMM.CloseImage Tile67
            
        ElseIf i = 68 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile68", Tile68
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile68, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile68, "D:\Users\JohnE\VisualBasicTests\Tile68.tif", False, 3

            'Close image
            PubMM.CloseImage Tile68
            
        ElseIf i = 69 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile69", Tile69
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile69, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile69, "D:\Users\JohnE\VisualBasicTests\Tile69.tif", False, 3

            'Close image
            PubMM.CloseImage Tile69

        ElseIf i = 70 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile70", Tile70
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile70, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile70, "D:\Users\JohnE\VisualBasicTests\Tile70.tif", False, 3

            'Close image
            PubMM.CloseImage Tile70
            
        ElseIf i = 71 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile71", Tile71
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile71, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile71, "D:\Users\JohnE\VisualBasicTests\Tile71.tif", False, 3

            'Close image
            PubMM.CloseImage Tile71

        ElseIf i = 72 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile72", Tile72
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile72, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile72, "D:\Users\JohnE\VisualBasicTests\Tile72.tif", False, 3

            'Close image
            PubMM.CloseImage Tile72
        
        ElseIf i = 73 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile73", Tile73
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile73, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile73, "D:\Users\JohnE\VisualBasicTests\Tile73.tif", False, 3

            'Close image
            PubMM.CloseImage Tile73
            
        ElseIf i = 74 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile74", Tile74
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile74, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile74, "D:\Users\JohnE\VisualBasicTests\Tile74.tif", False, 3

            'Close image
            PubMM.CloseImage Tile74
            
        ElseIf i = 75 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile75", Tile75
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile75, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile75, "D:\Users\JohnE\VisualBasicTests\Tile75.tif", False, 3

            'Close image
            PubMM.CloseImage Tile75

        ElseIf i = 76 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile76", Tile76
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile76, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile76, "D:\Users\JohnE\VisualBasicTests\Tile76.tif", False, 3

            'Close image
            PubMM.CloseImage Tile76

        ElseIf i = 77 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile77", Tile77
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile77, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile77, "D:\Users\JohnE\VisualBasicTests\Tile77.tif", False, 3

            'Close image
            PubMM.CloseImage Tile77
            
        ElseIf i = 78 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile78", Tile78
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile78, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile78, "D:\Users\JohnE\VisualBasicTests\Tile78.tif", False, 3

            'Close image
            PubMM.CloseImage Tile78
        
        ElseIf i = 79 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile79", Tile79
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile79, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile79, "D:\Users\JohnE\VisualBasicTests\Tile79.tif", False, 3

            'Close image
            PubMM.CloseImage Tile79
            
        ElseIf i = 80 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile80", Tile80
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile80, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile80, "D:\Users\JohnE\VisualBasicTests\Tile80.tif", False, 3

            'Close image
            PubMM.CloseImage Tile80
            
        ElseIf i = 81 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile81", Tile81
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile81, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile81, "D:\Users\JohnE\VisualBasicTests\Tile81.tif", False, 3

            'Close image
            PubMM.CloseImage Tile81
            
        ElseIf i = 82 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile82", Tile82
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile82, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile82, "D:\Users\JohnE\VisualBasicTests\Tile82.tif", False, 3

            'Close image
            PubMM.CloseImage Tile82
            
        ElseIf i = 83 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile83", Tile83
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile83, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile83, "D:\Users\JohnE\VisualBasicTests\Tile83.tif", False, 3

            'Close image
            PubMM.CloseImage Tile83
            
        ElseIf i = 84 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile84", Tile84
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile84, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile84, "D:\Users\JohnE\VisualBasicTests\Tile84.tif", False, 3

            'Close image
            PubMM.CloseImage Tile84
            
        ElseIf i = 85 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile85", Tile85
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile85, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile85, "D:\Users\JohnE\VisualBasicTests\Tile85.tif", False, 3

            'Close image
            PubMM.CloseImage Tile85
            
        ElseIf i = 86 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile86", Tile86
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile86, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile86, "D:\Users\JohnE\VisualBasicTests\Tile86.tif", False, 3

            'Close image
            PubMM.CloseImage Tile86
            
        ElseIf i = 87 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile87", Tile87
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile87, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile87, "D:\Users\JohnE\VisualBasicTests\Tile87.tif", False, 3

            'Close image
            PubMM.CloseImage Tile87
            
        ElseIf i = 88 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile88", Tile88
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile88, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile88, "D:\Users\JohnE\VisualBasicTests\Tile88.tif", False, 3

            'Close image
            PubMM.CloseImage Tile88
            
        ElseIf i = 89 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile89", Tile89
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile89, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile89, "D:\Users\JohnE\VisualBasicTests\Tile89.tif", False, 3

            'Close image
            PubMM.CloseImage Tile89
            
        ElseIf i = 90 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile90", Tile90
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile90, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile90, "D:\Users\JohnE\VisualBasicTests\Tile90.tif", False, 3

            'Close image
            PubMM.CloseImage Tile90
            
        ElseIf i = 91 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile91", Tile91
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile91, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile91, "D:\Users\JohnE\VisualBasicTests\Tile91.tif", False, 3

            'Close image
            PubMM.CloseImage Tile91
            
         ElseIf i = 92 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile92", Tile92
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile92, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile92, "D:\Users\JohnE\VisualBasicTests\Tile92.tif", False, 3

            'Close image
            PubMM.CloseImage Tile92
            
        ElseIf i = 93 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile93", Tile93
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile93, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile93, "D:\Users\JohnE\VisualBasicTests\Tile93.tif", False, 3

            'Close image
            PubMM.CloseImage Tile93
        
        ElseIf i = 94 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile94", Tile94
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile94, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile94, "D:\Users\JohnE\VisualBasicTests\Tile94.tif", False, 3

            'Close image
            PubMM.CloseImage Tile94

        ElseIf i = 95 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile95", Tile95
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile95, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile95, "D:\Users\JohnE\VisualBasicTests\Tile95.tif", False, 3

            'Close image
            PubMM.CloseImage Tile95

        ElseIf i = 96 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile96", Tile96
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile96, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile96, "D:\Users\JohnE\VisualBasicTests\Tile96.tif", False, 3

            'Close image
            PubMM.CloseImage Tile96

        ElseIf i = 97 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile97", Tile97
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile97, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile97, "D:\Users\JohnE\VisualBasicTests\Tile97.tif", False, 3

            'Close image
            PubMM.CloseImage Tile97

        ElseIf i = 98 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile98", Tile98
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile98, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile98, "D:\Users\JohnE\VisualBasicTests\Tile98.tif", False, 3

            'Close image
            PubMM.CloseImage Tile98
    
        ElseIf i = 99 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile99", Tile99
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile99, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile99, "D:\Users\JohnE\VisualBasicTests\Tile99.tif", False, 3

            'Close image
            PubMM.CloseImage Tile99

        ElseIf i = 100 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile100", Tile100
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile100, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile100, "D:\Users\JohnE\VisualBasicTests\Tile100.tif", False, 3

            'Close image
            PubMM.CloseImage Tile100

        ElseIf i = 101 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile101", Tile101
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile101, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile101, "D:\Users\JohnE\VisualBasicTests\Tile101.tif", False, 3

            'Close image
            PubMM.CloseImage Tile101

        ElseIf i = 102 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile102", Tile102
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile102, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile102, "D:\Users\JohnE\VisualBasicTests\Tile102.tif", False, 3

            'Close image
            PubMM.CloseImage Tile102

        ElseIf i = 103 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile103", Tile103
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile103, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile103, "D:\Users\JohnE\VisualBasicTests\Tile103.tif", False, 3

            'Close image
            PubMM.CloseImage Tile103

        ElseIf i = 104 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile104", Tile104
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile104, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile104, "D:\Users\JohnE\VisualBasicTests\Tile104.tif", False, 3

            'Close image
            PubMM.CloseImage Tile104
        
        ElseIf i = 105 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile105", Tile105
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile105, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile105, "D:\Users\JohnE\VisualBasicTests\Tile105.tif", False, 3

            'Close image
            PubMM.CloseImage Tile105
            
        ElseIf i = 106 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile106", Tile106
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile106, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile106, "D:\Users\JohnE\VisualBasicTests\Tile106.tif", False, 3

            'Close image
            PubMM.CloseImage Tile106
            
        ElseIf i = 107 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile107", Tile107
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile107, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile107, "D:\Users\JohnE\VisualBasicTests\Tile107.tif", False, 3

            'Close image
            PubMM.CloseImage Tile107
           
         ElseIf i = 108 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile108", Tile108
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile108, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile108, "D:\Users\JohnE\VisualBasicTests\Tile108.tif", False, 3

            'Close image
            PubMM.CloseImage Tile108
            
        ElseIf i = 109 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile109", Tile109
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile109, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile109, "D:\Users\JohnE\VisualBasicTests\Tile109.tif", False, 3

            'Close image
            PubMM.CloseImage Tile109
            
        ElseIf i = 110 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile110", Tile110
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile110, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile110, "D:\Users\JohnE\VisualBasicTests\Tile110.tif", False, 3

            'Close image
            PubMM.CloseImage Tile110
            
        ElseIf i = 111 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile111", Tile111
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile111, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile111, "D:\Users\JohnE\VisualBasicTests\Tile111.tif", False, 3

            'Close image
            PubMM.CloseImage Tile111
            
        ElseIf i = 112 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile112", Tile112
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile112, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile112, "D:\Users\JohnE\VisualBasicTests\Tile112.tif", False, 3

            'Close image
            PubMM.CloseImage Tile112
            
        ElseIf i = 113 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile113", Tile113
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile113, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile113, "D:\Users\JohnE\VisualBasicTests\Tile113.tif", False, 3

            'Close image
            PubMM.CloseImage Tile113
            
        ElseIf i = 114 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile114", Tile114
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile114, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile114, "D:\Users\JohnE\VisualBasicTests\Tile114.tif", False, 3

            'Close image
            PubMM.CloseImage Tile114
            
        ElseIf i = 115 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile115", Tile115
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile115, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile115, "D:\Users\JohnE\VisualBasicTests\Tile115.tif", False, 3

            'Close image
            PubMM.CloseImage Tile115
        
        ElseIf i = 116 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile116", Tile116
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile116, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile116, "D:\Users\JohnE\VisualBasicTests\Tile116.tif", False, 3

            'Close image
            PubMM.CloseImage Tile116
            
        ElseIf i = 117 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile117", Tile117
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile117, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile117, "D:\Users\JohnE\VisualBasicTests\Tile117.tif", False, 3

            'Close image
            PubMM.CloseImage Tile117
            
        ElseIf i = 118 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile118", Tile118
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile118, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile118, "D:\Users\JohnE\VisualBasicTests\Tile118.tif", False, 3

            'Close image
            PubMM.CloseImage Tile118
            
        ElseIf i = 119 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile119", Tile119
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile119, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile119, "D:\Users\JohnE\VisualBasicTests\Tile119.tif", False, 3

            'Close image
            PubMM.CloseImage Tile119
            
        ElseIf i = 120 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile120", Tile120
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile120, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile120, "D:\Users\JohnE\VisualBasicTests\Tile120.tif", False, 3

            'Close image
            PubMM.CloseImage Tile120

        ElseIf i = 121 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile121", Tile121
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile121, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile121, "D:\Users\JohnE\VisualBasicTests\Tile121.tif", False, 3

            'Close image
            PubMM.CloseImage Tile121

        ElseIf i = 122 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile122", Tile122
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile122, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile122, "D:\Users\JohnE\VisualBasicTests\Tile122.tif", False, 3

            'Close image
            PubMM.CloseImage Tile122

        ElseIf i = 123 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile123", Tile123
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile123, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile123, "D:\Users\JohnE\VisualBasicTests\Tile123.tif", False, 3

            'Close image
            PubMM.CloseImage Tile123

        ElseIf i = 124 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile124", Tile124
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile124, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile124, "D:\Users\JohnE\VisualBasicTests\Tile124.tif", False, 3

            'Close image
            PubMM.CloseImage Tile124

        ElseIf i = 125 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile125", Tile125
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile125, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile125, "D:\Users\JohnE\VisualBasicTests\Tile125.tif", False, 3

            'Close image
            PubMM.CloseImage Tile125

        ElseIf i = 126 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile126", Tile126
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile126, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile126, "D:\Users\JohnE\VisualBasicTests\Tile126.tif", False, 3

            'Close image
            PubMM.CloseImage Tile126

        ElseIf i = 127 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile127", Tile127
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile127, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile127, "D:\Users\JohnE\VisualBasicTests\Tile127.tif", False, 3

            'Close image
            PubMM.CloseImage Tile127

        ElseIf i = 128 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile128", Tile128
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile128, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile128, "D:\Users\JohnE\VisualBasicTests\Tile128.tif", False, 3

            'Close image
            PubMM.CloseImage Tile128
            
        ElseIf i = 129 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile129", Tile129
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile129, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile129, "D:\Users\JohnE\VisualBasicTests\Tile129.tif", False, 3

            'Close image
            PubMM.CloseImage Tile129

        ElseIf i = 130 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile130", Tile130
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile130, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile130, "D:\Users\JohnE\VisualBasicTests\Tile130.tif", False, 3

            'Close image
            PubMM.CloseImage Tile130

        ElseIf i = 131 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile131", Tile131
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile131, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile131, "D:\Users\JohnE\VisualBasicTests\Tile131.tif", False, 3

            'Close image
            PubMM.CloseImage Tile131
            
         ElseIf i = 132 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile132", Tile132
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile132, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile132, "D:\Users\JohnE\VisualBasicTests\Tile132.tif", False, 3

            'Close image
            PubMM.CloseImage Tile132
            
        ElseIf i = 133 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile133", Tile133
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile133, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile133, "D:\Users\JohnE\VisualBasicTests\Tile133.tif", False, 3

            'Close image
            PubMM.CloseImage Tile133
            
        ElseIf i = 134 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile134", Tile134
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile134, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile134, "D:\Users\JohnE\VisualBasicTests\Tile134.tif", False, 3

            'Close image
            PubMM.CloseImage Tile134
            
        ElseIf i = 135 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile135", Tile135
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile135, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile135, "D:\Users\JohnE\VisualBasicTests\Tile135.tif", False, 3

            'Close image
            PubMM.CloseImage Tile135
            
        ElseIf i = 136 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile136", Tile136
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile136, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile136, "D:\Users\JohnE\VisualBasicTests\Tile136.tif", False, 3

            'Close image
            PubMM.CloseImage Tile136
            
        ElseIf i = 137 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile137", Tile137
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile137, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile137, "D:\Users\JohnE\VisualBasicTests\Tile137.tif", False, 3

            'Close image
            PubMM.CloseImage Tile137

        ElseIf i = 138 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile138", Tile138
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile138, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile138, "D:\Users\JohnE\VisualBasicTests\Tile138.tif", False, 3

            'Close image
            PubMM.CloseImage Tile138

        ElseIf i = 139 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile139", Tile139
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile139, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile139, "D:\Users\JohnE\VisualBasicTests\Tile139.tif", False, 3

            'Close image
            PubMM.CloseImage Tile139

        ElseIf i = 140 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile140", Tile140
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile140, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile140, "D:\Users\JohnE\VisualBasicTests\Tile140.tif", False, 3

            'Close image
            PubMM.CloseImage Tile140

        ElseIf i = 141 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile141", Tile141
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile141, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile141, "D:\Users\JohnE\VisualBasicTests\Tile141.tif", False, 3

            'Close image
            PubMM.CloseImage Tile141

        ElseIf i = 142 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile142", Tile142
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile142, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile142, "D:\Users\JohnE\VisualBasicTests\Tile142.tif", False, 3

            'Close image
            PubMM.CloseImage Tile142
    
        ElseIf i = 143 And TilesToFire(i) = 1 Then
            PubMM.CreateImage 512, 512, 16, "Tile143", Tile143
            
            'Creating image
            For j = 0 To 511
            For k = 0 To 511
                PubMM.ReadPixel BigMask, j + xScale2, k + yScale2, Pix
                PubMM.WritePixel Tile143, j, k, Pix
            Next k
            Next j
            
            'saving
            PubMM.SaveImage Tile143, "D:\Users\JohnE\VisualBasicTests\Tile143.tif", False, 3

            'Close image
            PubMM.CloseImage Tile143


        End If
        
        
       
    Next i


    'Close image
    PubMM.CloseImage BigMask


End If


End Sub



Private Sub Command18_Click()


'Initializations
Dim Tile0a As Long
Dim Tile1a As Long
Dim Tile2a As Long
Dim Tile3a As Long
Dim Tile4a As Long
Dim Tile5a As Long
Dim Tile6a As Long
Dim Tile7a As Long
Dim Tile8a As Long
Dim Tile9a As Long
Dim Tile10a As Long
Dim Tile11a As Long
Dim Tile12a As Long
Dim Tile13a As Long
Dim Tile14a As Long
Dim Tile15a As Long
Dim Tile16a As Long
Dim Tile17a As Long
Dim Tile18a As Long
Dim Tile19a As Long
Dim Tile20a As Long
Dim Tile21a As Long
Dim Tile22a As Long
Dim Tile23a As Long
Dim Tile24a As Long
Dim Tile25a As Long
Dim Tile26a As Long
Dim Tile27a As Long
Dim Tile28a As Long
Dim Tile29a As Long
Dim Tile30a As Long
Dim Tile31a As Long
Dim Tile32a As Long
Dim Tile33a As Long
Dim Tile34a As Long
Dim Tile35a As Long
Dim Tile36a As Long
Dim Tile37a As Long
Dim Tile38a As Long
Dim Tile39a As Long
Dim Tile40a As Long
Dim Tile41a As Long
Dim Tile42a As Long
Dim Tile43a As Long
Dim Tile44a As Long
Dim Tile45a As Long
Dim Tile46a As Long
Dim Tile47a As Long
Dim Tile48a As Long
Dim Tile49a As Long
Dim Tile50a As Long
Dim Tile51a As Long
Dim Tile52a As Long
Dim Tile53a As Long
Dim Tile54a As Long
Dim Tile55a As Long
Dim Tile56a As Long
Dim Tile57a As Long
Dim Tile58a As Long
Dim Tile59a As Long
Dim Tile60a As Long
Dim Tile61a As Long
Dim Tile62a As Long
Dim Tile63a As Long
Dim Tile64a As Long
Dim Tile65a As Long
Dim Tile66a As Long
Dim Tile67a As Long
Dim Tile68a As Long
Dim Tile69a As Long
Dim Tile70a As Long
Dim Tile71a As Long
Dim Tile72a As Long
Dim Tile73a As Long
Dim Tile74a As Long
Dim Tile75a As Long
Dim Tile76a As Long
Dim Tile77a As Long
Dim Tile78a As Long
Dim Tile79a As Long
Dim Tile80a As Long
Dim Tile81a As Long
Dim Tile82a As Long
Dim Tile83a As Long
Dim Tile84a As Long
Dim Tile85a As Long
Dim Tile86a As Long
Dim Tile87a As Long
Dim Tile88a As Long
Dim Tile89a As Long
Dim Tile90a As Long
Dim Tile91a As Long
Dim Tile92a As Long
Dim Tile93a As Long
Dim Tile94a As Long
Dim Tile95a As Long
Dim Tile96a As Long
Dim Tile97a As Long
Dim Tile98a As Long
Dim Tile99a As Long
Dim Tile100a As Long
Dim Tile101a As Long
Dim Tile102a As Long
Dim Tile103a As Long
Dim Tile104a As Long
Dim Tile105a As Long
Dim Tile106a As Long
Dim Tile107a As Long
Dim Tile108a As Long
Dim Tile109a As Long
Dim Tile110a As Long
Dim Tile111a As Long
Dim Tile112a As Long
Dim Tile113a As Long
Dim Tile114a As Long
Dim Tile115a As Long
Dim Tile116a As Long
Dim Tile117a As Long
Dim Tile118a As Long
Dim Tile119a As Long
Dim Tile120a As Long
Dim Tile121a As Long
Dim Tile122a As Long
Dim Tile123a As Long
Dim Tile124a As Long
Dim Tile125a As Long
Dim Tile126a As Long
Dim Tile127a As Long
Dim Tile128a As Long
Dim Tile129a As Long
Dim Tile130a As Long
Dim Tile131a As Long
Dim Tile132a As Long
Dim Tile133a As Long
Dim Tile134a As Long
Dim Tile135a As Long
Dim Tile136a As Long
Dim Tile137a As Long
Dim Tile138a As Long
Dim Tile139a As Long
Dim Tile140a As Long
Dim Tile141a As Long
Dim Tile142a As Long
Dim Tile143a As Long
Dim CurrImMorgan As Long


'Changing the Dichroic
Dim Msg As String
Dim Title As String
Dim Help As String
Dim Ctxt As Integer
Dim Style As VbMsgBoxStyle
Dim theResult As VbMsgBoxResult

'Move to the DAPI Cube
If Option11.Value = True Then

    'move to the dapi cube
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_mosaic_channel.jnl"

    Title = "Check FL Shutter"
    Msg = "Hit FL on/off and press Yes"
    Help = "DEMO.HLP"
    Ctxt = 1000
    Style = vbYesNo

    theResult = MsgBox(Msg, Style, Title, Help, Ctxt)
    
'Move to the Split TIRF Cube
ElseIf Option13.Value = True Then
    
    'move to the split TIRF Cube
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_split_tirf_channel_mosaic.jnl"
    
    Title = "Check FL Shutter"
    Msg = "Hit FL on/off and press Yes"
    Help = "DEMO.HLP"
    Ctxt = 1000
    Style = vbYesNo

    theResult = MsgBox(Msg, Style, Title, Help, Ctxt)
    

End If

'Firing the tiles
If TilesToFire(0) = 1 Then
    
    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile0.tif", Tile0a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(0)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(0)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile0a
    
End If
    
If TilesToFire(1) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile1.tif", Tile1a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(1)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(1)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile1a
    
End If
    
If TilesToFire(2) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile2.tif", Tile2a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(2)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(2)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile2a
   
End If

If TilesToFire(3) = 1 Then
   
    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile3.tif", Tile3a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(3)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(3)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile3a
    
End If

If TilesToFire(4) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile4.tif", Tile4a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(4)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(4)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile4a
    
End If

If TilesToFire(5) = 1 Then
    
    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile5.tif", Tile5a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(5)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(5)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile5a
    
End If

If TilesToFire(6) = 1 Then
    
    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile6.tif", Tile6a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(6)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(6)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile6a
    
End If

If TilesToFire(7) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile7.tif", Tile7a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(7)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(7)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile7a
    
End If

If TilesToFire(8) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile8.tif", Tile8a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(8)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(8)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile8a
    
End If

If TilesToFire(9) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile9.tif", Tile9a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(9)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(9)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile9a
    
End If

If TilesToFire(10) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile10.tif", Tile10a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(10)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(10)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile10a
    
End If

If TilesToFire(11) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile11.tif", Tile11a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(11)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(11)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile11a
    
End If

If TilesToFire(12) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile12.tif", Tile12a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(12)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(12)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile12a
    
End If

If TilesToFire(13) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile13.tif", Tile13a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(13)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(13)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile13a
    
End If

If TilesToFire(14) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile14.tif", Tile14a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(14)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(14)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile14a
    
End If

If TilesToFire(15) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile15.tif", Tile15a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(15)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(15)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile15a
    
End If

If TilesToFire(16) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile16.tif", Tile16a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(16)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(16)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile16a
    
End If

If TilesToFire(17) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile17.tif", Tile17a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(17)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(17)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile17a
    
End If

If TilesToFire(18) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile18.tif", Tile18a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(18)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(18)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile18a
    
End If

If TilesToFire(19) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile19.tif", Tile19a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(19)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(19)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile19a
    
End If

If TilesToFire(20) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile20.tif", Tile20a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(20)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(20)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile20a
    
End If

If TilesToFire(21) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile21.tif", Tile21a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(21)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(21)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile21a
    
End If

If TilesToFire(22) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile22.tif", Tile22a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(22)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(22)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile22a
    
End If

If TilesToFire(23) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile23.tif", Tile23a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(23)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(23)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile23a
    
End If

If TilesToFire(24) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile24.tif", Tile24a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(24)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(24)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile24a
    
End If

If TilesToFire(25) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile25.tif", Tile25a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(25)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(25)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile25a
    
End If

If TilesToFire(26) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile26.tif", Tile26a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(26)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(26)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile26a
    
End If

If TilesToFire(27) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile27.tif", Tile27a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(27)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(27)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile27a
    
End If

If TilesToFire(28) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile28.tif", Tile28a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(28)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(28)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile28a
    
End If

If TilesToFire(29) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile29.tif", Tile29a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(29)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(29)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile29a
    
End If

If TilesToFire(30) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile30.tif", Tile30a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(30)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(30)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile30a
    
End If

If TilesToFire(31) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile31.tif", Tile31a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(31)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(31)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile31a
    
End If

If TilesToFire(32) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile32.tif", Tile32a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(32)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(32)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile32a
    
End If

If TilesToFire(33) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile33.tif", Tile33a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(33)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(33)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile33a
    
End If

If TilesToFire(34) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile34.tif", Tile34a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(34)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(34)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile34a
    
End If

If TilesToFire(35) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile35.tif", Tile35a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(35)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(35)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile35a
    
End If

If TilesToFire(36) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile36.tif", Tile36a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(36)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(36)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile36a
    
End If

If TilesToFire(37) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile37.tif", Tile37a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(37)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(37)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile37a
    
End If

If TilesToFire(38) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile38.tif", Tile38a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(38)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(38)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile38a
    
End If

If TilesToFire(39) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile39.tif", Tile39a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(39)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(39)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile39a
    
End If

If TilesToFire(40) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile40.tif", Tile40a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(40)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(40)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile40a
    
End If

If TilesToFire(41) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile41.tif", Tile41a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(41)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(41)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile41a
    
End If

If TilesToFire(42) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile42.tif", Tile42a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(42)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(42)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile42a
    
End If

If TilesToFire(43) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile43.tif", Tile43a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(43)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(43)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile43a
    
End If

If TilesToFire(44) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile44.tif", Tile44a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(44)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(44)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile44a
    
End If

If TilesToFire(45) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile45.tif", Tile45a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(45)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(45)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile45a
    
End If


If TilesToFire(46) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile46.tif", Tile46a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(46)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(46)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile46a
    
End If


If TilesToFire(47) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile47.tif", Tile47a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(47)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(47)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile47a
    
End If


If TilesToFire(48) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile48.tif", Tile48a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(48)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(48)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile48a
    
End If

If TilesToFire(49) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile49.tif", Tile49a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(49)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(49)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile49a
    
End If

If TilesToFire(50) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile50.tif", Tile50a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(50)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(50)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile50a
    
End If

If TilesToFire(51) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile51.tif", Tile51a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(51)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(51)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile51a
    
End If

If TilesToFire(52) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile52.tif", Tile52a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(52)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(52)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile52a
    
End If

If TilesToFire(53) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile53.tif", Tile53a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(53)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(53)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile53a
    
End If

If TilesToFire(54) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile54.tif", Tile54a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(54)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(54)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile54a
    
End If

If TilesToFire(55) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile55.tif", Tile55a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(55)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(55)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile55a
    
End If


If TilesToFire(56) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile56.tif", Tile56a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(56)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(56)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile56a
    
End If

If TilesToFire(57) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile57.tif", Tile57a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(57)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(57)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile57a
    
End If

If TilesToFire(58) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile58.tif", Tile58a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(58)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(58)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile58a
    
End If

If TilesToFire(59) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile59.tif", Tile59a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(59)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(59)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile59a
    
End If

If TilesToFire(60) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile60.tif", Tile60a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(60)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(60)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile60a
    
End If

If TilesToFire(61) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile61.tif", Tile61a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(61)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(61)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile61a
    
End If

If TilesToFire(62) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile62.tif", Tile62a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(62)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(62)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile62a
    
End If

If TilesToFire(63) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile63.tif", Tile63a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(63)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(63)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile63a
    
End If

If TilesToFire(64) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile64.tif", Tile64a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(64)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(64)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile64a
    
End If

If TilesToFire(65) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile65.tif", Tile65a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(65)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(65)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile65a
    
End If

If TilesToFire(66) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile66.tif", Tile66a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(66)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(66)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile66a
    
End If

If TilesToFire(67) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile67.tif", Tile67a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(67)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(67)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile67a
    
End If

If TilesToFire(68) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile68.tif", Tile68a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(68)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(68)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile68a
    
End If

If TilesToFire(69) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile69.tif", Tile69a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(69)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(69)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile69a
    
End If

If TilesToFire(70) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile70.tif", Tile70a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(70)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(70)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile70a
    
End If

If TilesToFire(71) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile71.tif", Tile71a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(71)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(71)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile71a
    
End If

If TilesToFire(72) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile72.tif", Tile72a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(72)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(72)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile72a
    
End If

If TilesToFire(73) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile73.tif", Tile73a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(73)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(73)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile73a
    
End If

If TilesToFire(74) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile74.tif", Tile74a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(74)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(74)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile74a
    
End If

If TilesToFire(75) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile75.tif", Tile75a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(75)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(75)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile75a
    
End If

If TilesToFire(76) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile76.tif", Tile76a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(76)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(76)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile76a
    
End If

If TilesToFire(77) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile77.tif", Tile77a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(77)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(77)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile77a
    
End If

If TilesToFire(78) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile78.tif", Tile78a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(78)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(78)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile78a
    
End If

If TilesToFire(79) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile79.tif", Tile79a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(79)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(79)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile79a
    
End If

If TilesToFire(80) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile80.tif", Tile80a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(80)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(80)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile80a
    
End If

If TilesToFire(81) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile81.tif", Tile81a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(81)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(81)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile81a
    
End If

If TilesToFire(82) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile82.tif", Tile82a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(82)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(82)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile82a
    
End If

If TilesToFire(83) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile83.tif", Tile83a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(83)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(83)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile83a
    
End If

If TilesToFire(84) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile84.tif", Tile84a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(84)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(84)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile84a
    
End If

If TilesToFire(85) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile85.tif", Tile85a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(85)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(85)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile85a
    
End If

If TilesToFire(86) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile86.tif", Tile86a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(86)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(86)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile86a
    
End If

If TilesToFire(87) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile87.tif", Tile87a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(87)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(87)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile87a
    
End If

If TilesToFire(88) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile88.tif", Tile88a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(88)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(88)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile88a
    
End If

If TilesToFire(89) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile89.tif", Tile89a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(89)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(89)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile89a
    
End If


If TilesToFire(90) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile90.tif", Tile90a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(90)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(90)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile90a
    
End If

If TilesToFire(91) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile91.tif", Tile91a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(91)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(91)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile91a
    
End If

If TilesToFire(92) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile92.tif", Tile92a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(92)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(92)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile92a
    
End If


If TilesToFire(93) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile93.tif", Tile93a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(93)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(93)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile93a
    
End If

If TilesToFire(94) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile94.tif", Tile94a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(94)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(94)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile94a
    
End If

If TilesToFire(95) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile95.tif", Tile95a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(95)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(95)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile95a
    
End If

If TilesToFire(96) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile96.tif", Tile96a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(96)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(96)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile96a
    
End If

If TilesToFire(97) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile97.tif", Tile97a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(97)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(97)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile97a
    
End If

If TilesToFire(98) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile98.tif", Tile98a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(98)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(98)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile98a
    
End If

If TilesToFire(99) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile99.tif", Tile99a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(99)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(99)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile99a
    
End If

If TilesToFire(100) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile100.tif", Tile100a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(100)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(100)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile100a
    
End If

If TilesToFire(101) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile101.tif", Tile101a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(101)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(101)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile101a
    
End If

If TilesToFire(102) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile102.tif", Tile102a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(102)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(102)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile102a
    
End If

If TilesToFire(103) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile103.tif", Tile103a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(103)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(103)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile103a
    
End If

If TilesToFire(104) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile104.tif", Tile104a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(104)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(104)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile104a
    
End If

If TilesToFire(105) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile105.tif", Tile105a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(105)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(105)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile105a
    
End If

If TilesToFire(106) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile106.tif", Tile106a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(106)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(106)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile106a
    
End If

If TilesToFire(107) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile107.tif", Tile107a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(107)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(107)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile107a
    
End If

If TilesToFire(108) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile108.tif", Tile108a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(108)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(108)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile108a
    
End If

If TilesToFire(109) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile109.tif", Tile109a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(109)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(109)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile109a
    
End If

If TilesToFire(110) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile110.tif", Tile110a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(110)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(110)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile110a
    
End If

If TilesToFire(111) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile111.tif", Tile111a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(111)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(111)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile111a
    
End If

If TilesToFire(112) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile112.tif", Tile112a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(112)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(112)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile112a
    
End If


If TilesToFire(113) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile113.tif", Tile113a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(113)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(113)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile113a
    
End If


If TilesToFire(114) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile114.tif", Tile114a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(114)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(114)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile114a
    
End If

If TilesToFire(115) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile115.tif", Tile115a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(115)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(115)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile115a
    
End If

If TilesToFire(116) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile116.tif", Tile116a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(116)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(116)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile116a
    
End If

If TilesToFire(117) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile117.tif", Tile117a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(117)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(117)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile117a
    
End If

If TilesToFire(118) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile118.tif", Tile118a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(118)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(118)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile118a
    
End If

If TilesToFire(119) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile119.tif", Tile119a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(119)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(119)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile119a
    
End If

If TilesToFire(120) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile120.tif", Tile120a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(120)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(120)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile120a
    
End If

If TilesToFire(121) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile121.tif", Tile121a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(121)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(121)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile121a
    
End If

If TilesToFire(122) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile122.tif", Tile122a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(122)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(122)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile122a
    
End If

If TilesToFire(123) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile123.tif", Tile123a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(123)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(123)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile123a
    
End If

If TilesToFire(124) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile124.tif", Tile124a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(124)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(124)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile124a
    
End If

If TilesToFire(125) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile125.tif", Tile125a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(125)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(125)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile125a
    
End If

If TilesToFire(126) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile126.tif", Tile126a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(126)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(126)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile126a
    
End If

If TilesToFire(127) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile127.tif", Tile127a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(127)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(127)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile127a
    
End If

If TilesToFire(128) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile128.tif", Tile128a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(128)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(128)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile128a
    
End If

If TilesToFire(129) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile129.tif", Tile129a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(129)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(129)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile129a
    
End If

If TilesToFire(130) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile130.tif", Tile130a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(130)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(130)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile130a
    
End If

If TilesToFire(131) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile131.tif", Tile131a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(131)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(131)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile131a
    
End If

If TilesToFire(132) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile132.tif", Tile132a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(132)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(132)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile132a
    
End If

If TilesToFire(133) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile133.tif", Tile133a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(133)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(133)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile133a
    
End If

If TilesToFire(134) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile134.tif", Tile134a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(134)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(134)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile134a
    
End If

If TilesToFire(135) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile135.tif", Tile135a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(135)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(135)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile135a
    
End If

If TilesToFire(136) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile136.tif", Tile136a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(136)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(136)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile136a
    
End If

If TilesToFire(137) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile137.tif", Tile137a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(137)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(137)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile137a
    
End If

If TilesToFire(138) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile138.tif", Tile138a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(138)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(138)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile138a
    
End If

If TilesToFire(139) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile139.tif", Tile139a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(139)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(139)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile139a
    
End If

If TilesToFire(140) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile140.tif", Tile140a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(140)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(140)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile140a
    
End If

If TilesToFire(141) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile141.tif", Tile141a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(141)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(141)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile141a
    
End If

If TilesToFire(142) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile142.tif", Tile142a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(142)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(142)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile142a
    
End If

If TilesToFire(143) = 1 Then

    'loading
    PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\Tile143.tif", Tile143a
    
    'Move Stage
    PubMM.SetMMVariable "Device.Stage.XPosition", xTilePosGlobal(143)
    PubMM.SetMMVariable "Device.Stage.YPosition", yTilePosGlobal(143)
    
    'Fire Mosaic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire2.jnl"
    
    'Get the current image
    PubMM.GetCurrentImage CurrImMorgan

    'This is a line to save image to dummy directory so that you do not have to keep hitting  "No I do not want to save" as you tile
    PubMM.SaveImage CurrImMorgan, "D:\Users\JohnE\tmp\tmpImageMask.tif", False, 3
    
    'Close image
    PubMM.CloseImage Tile143a
    
End If

'If user selected the DAPI cube, I am swinging to the TIRF cube back in
If Option11.Value = True Or Option13.Value = True Then

    'swinging in the multi-color tirf dichroic
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_tirf_channel.jnl"

    'Message box to manually hit shutter
    Title = "Check FL Shutter"
    Msg = "Hit FL on/off and press Yes"
    Help = "DEMO.HLP"
    Ctxt = 1000
    Style = vbYesNo

    theResult = MsgBox(Msg, Style, Title, Help, Ctxt)

End If


'Close the big image
'PubMM.CloseImage BigMask
End Sub

Private Sub Command19_Click()

'This is the button to set the refill rate

'Definitions
Dim Rate As Single
Dim RateString As String
Dim StringSend2 As String


'Getting the input from the text box
RateString = Text12.Text

'Making floating point
InRate = CSng(RateString)

'Trying this with a string
StringSend2 = "RFR " + RateString + "MM"

'Changing the Volume
MSComm1.Output = StringSend2 + Chr(13)

End Sub

Private Sub Command2_Click()

'fire mosiac
PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to fire mosaic\mosaic_fire.jnl"

  

End Sub

Private Sub Command20_Click()


        'Decrement ROI Number
        ROINum = ROINum - 1
        
        'putting a picture in gui
        Picture1.Picture = LoadPicture("D:\Users\JohnE\VisualBasicTests\TileTest2.bmp")

        'forcing the image to fit in the box
        Picture1.ScaleMode = 3
        Picture1.AutoRedraw = True
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight

        ' draw previous ROIs
        For u = 1 To ROINum
            For s = 0 To (MasterROICounter + 1)
                    If IdxAllROIs(s + 1) = u And IdxAllROIs(s) = u Then
                        Picture1.Line (xDrawAllROIs(s), yDrawAllROIs(s))-(xDrawAllROIs(s + 1), yDrawAllROIs(s + 1)), vbRed
                    End If
            Next s
        Next u

        
        'initializing the counter in which I keep track of drawing of coordinates
        theCounter = 0
    
        'initializing the arrays that hold xy coordinates of drawing
        For i = 0 To 99
            xDraw(i) = 0.1
            yDraw(i) = 0.1
        Next i
                
        'Removing unwanted ROI from master list
        For r = 0 To 999
            If IdxAllROIs(r) > ROINum Then
                xDrawAllROIs(r) = 0
                yDrawAllROIs(r) = 0
                IdxAllROIs(r) = 0
                MasterROICounter = MasterROICounter - 1
            End If
         Next r


End Sub

Private Sub Command22_Click()

'This is a button to test if I can do file IO with text files
Dim tmp As String
Dim contents As String
Open "D:\Users\JohnE\VisualBasicTests\Contacts.txt" For Input As #1
While EOF(1) = 0
    Line Input #1, tmp
    contents = contents + tmp
Wend
Close #1
MsgBox contents

End Sub

Private Sub Command21_Click()

'This is the reset and make image brighter button


     'Reset ROI number
        ROINum = 0
        
        'Load the image into MetaMorph
        Dim theImNow As Long
        PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\TileTest2.bmp", theImNow
        
        'make it brighter
        If HScroll1.Value = 0 Then
            TheBrightness = TheBrightness * 1
        Else
            TheBrightness = TheBrightness * (HScroll1.Value + 1)
        End If
        PubMM.SetBrightness theImNow, TheBrightness
        PubMM.FixImage theImNow
        
        'saving the image
        'Saving the big tiled image as *.bmp
        Dim FileExt As Integer
        FileExt = 2
        PubMM.SaveImage theImNow, "D:\Users\JohnE\VisualBasicTests\TileTest2.bmp", False, FileExt
        
        'close the image
        PubMM.CloseImage theImNow
        
        'putting a picture in gui
        Picture1.Picture = LoadPicture("D:\Users\JohnE\VisualBasicTests\TileTest2.bmp")

        'forcing the image to fit in the box
        Picture1.ScaleMode = 3
        Picture1.AutoRedraw = True
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
        
        'initializing the counter in which I keep track of drawing of coordinates
        MasterROICounter = 0
    
        'initializing the arrays that hold xy coordinates of drawing
        For i = 0 To 99
            xDraw(i) = 0.1
            yDraw(i) = 0.1
        Next i
                
        'initializing all ROI vertices
        For r = 0 To 999
            xDrawAllROIs(r) = 0
            yDrawAllROIs(r) = 0
            IdxAllROIs(r) = 0
         Next r


End Sub

Private Sub Command23_Click()

'This is button where I am testing how to write to a file
Open "D:\Users\JohnE\VisualBasicTests\OutPutTest.txt" For Output As #1

For v = 0 To 4
    If v = 0 Then
        Print #1, "Does this print to the file?"
    ElseIf v = 1 Then
        Print #1, "Is this on second line?"
    ElseIf v = 2 Then
        Print #1, "Is this on third line?"
    ElseIf v = 3 Then
        Print #1, "Is this on fourth line?"
    ElseIf v = 4 Then
        Print #1, "Is this on fifth line?"
    End If
Next v

Close #1



End Sub

Private Sub Command24_Click()

'This is the button to load old tiled images


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''This is the first call to the dialog box to load text file'''''''''''''''''''''

'Initialization
Dim theFileText As String
Dim tmp1 As String
Dim CountNov As Integer
Dim CountNovGlobalTile As Integer
Dim CountNovXorY As Integer
Dim StoredSizeFromTextFile As Integer
Dim TilePosFromNovTxtFile As Double
Dim xPosOldImage As Double
Dim yPosOldImage As Double
Dim CountKen As Integer

'asking user to go to text file
'This String contains the path and file concatenated
CommonDialog2.DialogTitle = "Select the Text File Associated w/ Image"
CommonDialog2.ShowOpen
theFileText = CommonDialog2.FileName

'Now I have to extract all the relevant information from the text file

'Opening the file for reading
Open theFileText For Input As #1

'initialize counters
CountKen = 0
CountNov = 0
CountNovGlobalTile = 0
CountNovXorY = 1

While EOF(1) = 0

    'Get the line from the text file as string
    Line Input #1, tmp1
    CountKen = CountKen + 1
    
    'Figuring out what each line of text file is
    
    'Size of tiled image
    If CountNov = 0 Then
        StoredSizeFromTextFile = CInt(tmp1)
        If StoredSizeFromTextFile = 2 Then
            Option3.Value = True
            Option4.Value = False
            Option5.Value = False
            Option6.Value = False
            Option7.Value = False
        ElseIf StoredSizeFromTextFile = 4 Then
            Option3.Value = False
            Option4.Value = True
            Option5.Value = False
            Option6.Value = False
            Option7.Value = False
        ElseIf StoredSizeFromTextFile = 6 Then
            Option3.Value = False
            Option4.Value = False
            Option5.Value = True
            Option6.Value = False
            Option7.Value = False
        ElseIf StoredSizeFromTextFile = 8 Then
            Option3.Value = False
            Option4.Value = False
            Option5.Value = False
            Option6.Value = True
            Option7.Value = False
        ElseIf StoredSizeFromTextFile = 12 Then
            Option3.Value = False
            Option4.Value = False
            Option5.Value = False
            Option6.Value = False
            Option7.Value = True
        End If
        
    'x center of tiled image
    ElseIf CountNov = 1 Then
        Text7.Text = tmp1
        Text1.Text = tmp1
        xPosOldImage = CDbl(tmp1)
        PubMM.SetMMVariable "Device.Stage.XPosition", xPosOldImage
    'y center of tiled image
    ElseIf CountNov = 2 Then
        Text8.Text = tmp1
        Text3.Text = tmp1
        yPosOldImage = CDbl(tmp1)
        PubMM.SetMMVariable "Device.Stage.YPosition", yPosOldImage
    'Picture1.Width
    ElseIf CountNov = 3 Then
        Picture1.Width = CInt(tmp1)
        
    'Picture1.Height
    ElseIf CountNov = 4 Then
        Picture1.Height = CInt(tmp1)
        
    'Picture1.Left
    ElseIf CountNov = 5 Then
        Picture1.Left = CInt(tmp1)
        
    'Picture1.Top
    ElseIf CountNov = 6 Then
        Picture1.Top = CInt(tmp1)
        
    'Tile Positions
    Else
    
        'get a coordinate
        TilePosFromNovTxtFile = CDbl(tmp1)
    
        If CountNovXorY = 1 Then
        
            'load the coordinate
            xTilePosGlobal(CountNovGlobalTile) = TilePosFromNovTxtFile

        ElseIf CountNovXorY = 2 Then
        
            'load the coordinate
            yTilePosGlobal(CountNovGlobalTile) = TilePosFromNovTxtFile
        
            'iterate master tile counter
            CountNovGlobalTile = CountNovGlobalTile + 1
            
            'reset counter
            CountNovXorY = 0

    
        End If
        
        'iterate counter
       CountNovXorY = CountNovXorY + 1
    
    End If
    
    'iterate counter
    CountNov = CountNov + 1
    
Wend
Close #1

'debugging
Text4.Text = CStr(CountNovGlobalTile)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''This is the second call to dialog box to load images'''''''''''''''''''''''''''

'Initializations
Dim theFile1 As String
Dim FileExtNow As Integer
Dim UserIm As Long
FileExtNow = 2

'asking user to go to picture
'This String contains the path and file concatenated
CommonDialog2.DialogTitle = "Select the Image"
CommonDialog2.ShowOpen
theFile1 = CommonDialog2.FileName

'Reset ROI number
ROINum = 0

'Loading the image selected by user
PubMM.LoadImage theFile1, UserIm

'Saving image as *.tif
PubMM.SaveImage UserIm, "D:\Users\JohnE\VisualBasicTests\TileTest2.bmp", False, FileExtNow

'Close the image
PubMM.CloseImage UserIm
        
'putting a picture in gui
Picture1.Picture = LoadPicture("D:\Users\JohnE\VisualBasicTests\TileTest2.bmp")

'forcing the image to fit in the box
Picture1.ScaleMode = 3
Picture1.AutoRedraw = True
Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
        
'initializing the counter in which I keep track of drawing of coordinates
MasterROICounter = 0
    
'initializing the arrays that hold xy coordinates of drawing
For i = 0 To 99
    xDraw(i) = 0.1
    yDraw(i) = 0.1
Next i
                
'initializing all ROI vertices
For r = 0 To 999
    xDrawAllROIs(r) = 0
    yDrawAllROIs(r) = 0
    IdxAllROIs(r) = 0
Next r


End Sub

Private Sub Command3_Click()

    'These are the selections for the illumination combo box
   If Combo1.Text = "GFP 100%" Then
        'reset illumination setting
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_gfp_channel.jnl"
    ElseIf Combo1.Text = "DAPI 100%" Then
        'reset illumination setting
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_dapi_channel.jnl"
    ElseIf Combo1.Text = "Cy5 100%" Then
        'reset illumination setting
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_cy5_channel.jnl"
    ElseIf Combo1.Text = "TxRd 100%" Then
        PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to set channels\jour_set_txred_channel_100.jnl"
    End If

    'start live
    PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to start live\start_live.jnl"

End Sub

Private Sub Command4_Click()

'end live
PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal to end live\end_live2.jnl"

End Sub

Private Sub Command5_Click()

'This is the button for Refill

'Setting to refill
Dim RefillString As String
RefillString = "DIR REF"
MSComm1.Output = RefillString + Chr(13)

'Running the Syringe
MSComm1.Output = "RUN" + Chr(13)

End Sub

Private Sub Command6_Click()

'This is the button to select an xy position on an image

'Load an image
'PubMM.LoadImage "D:\Users\JohnE\TIRF Move Oct 2018\100418\561nm excitation\field_488nm_2.tif", 1

'Initial xy positions
'Dim xPosStart As Double
'Dim yPosStart As Double
'Dim xPosStartString As String
'Dim yPosStartString As String

'Asking the user to pick a position
Dim xPosLong As Double
Dim yPosLong As Double
Dim xPosString As String
Dim yPosString As String

'variables to see what exactly the x position of the stage is
Dim xCheckString As String
Dim xCheck As Double

'Getting the initial xy coordinates
'PubMM.GetMMVariable "Device.Stage.XPosition", xPosStart
'PubMM.GetMMVariable "Device.Stage.YPosition", yPosStart

'Making initial xy coordinates strings
'xPosStartString = CStr(xPosStart)
'yPosStartString = CStr(yPosStart)
'Text7.Text = xPosStartString
'Text8.Text = yPosStartString

'Getting the input from the text box
xPosString = Text1.Text
yPosString = Text3.Text

'Casting as floating point
xPosLong = CDbl(xPosString)
yPosLong = CDbl(yPosString)

'Have the person pick a point - initializing PickPoint Variable
'PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\journal for picking points\journal to pick points.jnl"

'Getting the xy coordinates
'PubMM.GetMMVariable "PickPoint.X", xPos
'PubMM.GetMMVariable "PickPoint.Y", yPos

PubMM.SetMMVariable "Device.Stage.XPosition", xPosLong
PubMM.SetMMVariable "Device.Stage.YPosition", yPosLong

'checking the stage x position
'PubMM.GetMMVariable "Device.Stage.XPosition", xCheck
'xCheckString = CStr(xCheck)
'Text4.Text = CStr(xCheckString)





End Sub

Private Sub Command7_Click()

Dim theImage As Long
Dim theImageByte As Byte
Dim xP As Long
Dim yP As Long
Dim X As Single
Dim Y As Single
Dim Button As Integer
Dim Shift As Integer
Dim LocalMousePosition As Long
Dim Xstr As String
Dim Ystr As String
Dim XStart As Single
Dim YStart As Single
Dim XOld As Single
Dim YOld As Single
Dim i As Integer
Dim lut(256) As Byte
Dim lutnum As Integer
Dim bitDepth As Integer
Dim bitDepthStr As String
Dim theFileNow As String





'Load Image
'PubMM.LoadImage "D:\Users\JohnE\TIRF Move Oct 2018\100418\561nm excitation\field_488nm_2.tif", 1

'Get the active image
PubMM.GetCurrentImage theImage

'asking user to go to picture
CommonDialog1.ShowSave
theFileNow = CommonDialog1.FileName




'Saving image as rbg *.bmp
Dim FileExt As Integer
FileExt = 2
PubMM.SaveImage theImage, theFileNow, False, FileExt


'Text4.Text = CStr(xP)

'putting a picture in gui
Picture1.Picture = LoadPicture(theFileNow)
'Picture1.Picture = LoadPicture("D:\Users\JohnE\TIRF Move Oct 2018\100418\561nm excitation\field_488nm_2.tif")

'keeping the dimensions of the image the same as they were originally
Picture1.AutoSize = True

'Changing the color table of the image
'Dim i As Integer
'Dim lut(256) As Byte

'For i = 0 To 255
'lut(i) = 255 - i
'Next i

'PubMM.SetLut im, 10, 0, 256, lut, lut, lut
'PubMM.SetLutModel im, 1

'trying to get the mouse position
'LocalMousePosition = MouseDown(Button, Shift, X, Y)




End Sub




Private Sub Command8_Click()
MSComm1.CommPort = 5
MSComm1.PortOpen = True
MSComm1.Settings = "2400,N,8,2"





End Sub

Private Sub Command9_Click()

Dim theIm As Long
Dim theFile1 As String

'Get the active image
PubMM.GetCurrentImage theIm

'asking user to go to picture
CommonDialog2.ShowSave
theFile1 = CommonDialog2.FileName


'Saving image as *.tif
Dim FileExt1 As Integer
FileExt1 = 3
PubMM.SaveImage theIm, theFile1, False, FileExt1



End Sub

Private Sub HScroll1_Change()

'This is the horizontal scroll bar for brightness
HScroll1.Min = 1
HScroll1.Max = 5




End Sub

'Private Sub MSComm1_OnComm()

'MSComm1.CommPort = 5

'End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim xPosNow As Double
    Dim yPosNow As Double
    Dim xPosNowString As String
    Dim yPosNowString As String
    
    Dim xPosNew As Double
    Dim yPosNew As Double
    
    Dim xPosNewStr As String
    Dim yPosNewStr As String

    Dim theWidth As Long
    Dim theWidthStr As String
    Dim theIncrWidth As Double
    Dim theIncrString As String
    
    Dim theHeight As Long
    Dim theHeightStr As String
    Dim theIncrHeight As Double
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''When you are navigating on image''''''''''''''''''''
    
    If Option1.Value = True Then
    
    
        'putting a picture in gui
        Picture1.Picture = LoadPicture("D:\Users\JohnE\VisualBasicTests\TileTest2.bmp")

        'forcing the image to fit in the box
        Picture1.ScaleMode = 3
        Picture1.AutoRedraw = True
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight


        'Getting the initil x position
        xPosNowString = Text7.Text
        yPosNowString = Text8.Text
        xPosNow = CDbl(xPosNowString)
        yPosNow = CDbl(yPosNowString)
    
        'picking points on image
        If Button <> 1 Then Exit Sub
        XStart = X
        YStart = Y
        DrawMode = 7
    
        'making string
        Xstr = CStr(X)
        Ystr = CStr(Y)
    
        'Text4.Text = Xstr
        'Text6.Text = Ystr

        'For 20x, x width in stage units = 400
        'For 20x y width in stage units = 400

        '4x2 mosaic - navigation
        If Option3.Value = True Then
        
            'Width of image in pixel (picture box units)
            theWidth = 528
            theHeight = 262
    
            'The increment
            theIncrWidth = 1600 / theWidth
            theIncrHeight = 800 / theHeight
    
            'Calculating new stage position
            xPosNew = (X * theIncrWidth) + xPosNow - 800
            yPosNew = 800 - ((Y * theIncrHeight) - yPosNow + 400)
            xPosNewStr = CStr(xPosNew)
            yPosNewStr = CStr(yPosNew)
    
    
            'set the new stage position
            PubMM.SetMMVariable "Device.Stage.XPosition", xPosNew
            PubMM.SetMMVariable "Device.Stage.YPosition", yPosNew
    
            'Adding to the text box
            Text4.Text = xPosNewStr
            Text6.Text = yPosNewStr
    
            'making a box
            Picture1.Line (X - 66.1, Y + 65.75)-(X + 66.1, Y + 65.75), vbRed
            Picture1.Line (X + 66.1, Y + 65.75)-(X + 66.1, Y - 65.75), vbRed
            Picture1.Line (X - 66.1, Y - 65.75)-(X + 66.1, Y - 65.75), vbRed
            Picture1.Line (X - 66.1, Y + 65.75)-(X - 66.1, Y - 65.75), vbRed
        
        End If
        
        '4x4 mosaic - navigation
        If Option4.Value = True Then
        
            'Width of image in pixel (picture box units)
            theWidth = 390
            theHeight = 390
    
            'The increment
            theIncrWidth = 1600 / theWidth
            theIncrHeight = 1600 / theHeight
    
            'Calculating new stage position
            xPosNew = (X * theIncrWidth) + xPosNow - 800
            yPosNew = 1200 - ((Y * theIncrHeight) - yPosNow + 400)
            xPosNewStr = CStr(xPosNew)
            yPosNewStr = CStr(yPosNew)
    
    
            'set the new stage position
            PubMM.SetMMVariable "Device.Stage.XPosition", xPosNew
            PubMM.SetMMVariable "Device.Stage.YPosition", yPosNew
    
            'Adding to the text box
            Text4.Text = xPosNewStr
            Text6.Text = yPosNewStr
    
            'making a box
            Picture1.Line (X - 48.75, Y + 48.75)-(X + 48.75, Y + 48.75), vbRed
            Picture1.Line (X + 48.75, Y + 48.75)-(X + 48.75, Y - 48.75), vbRed
            Picture1.Line (X - 48.75, Y - 48.75)-(X + 48.75, Y - 48.75), vbRed
            Picture1.Line (X - 48.75, Y + 48.75)-(X - 48.75, Y - 48.75), vbRed
        
        End If
        
        '6x6 mosaic - navigation
        If Option5.Value = True Then
        
            'Width of image in pixel (picture box units)
            theWidth = 390
            theHeight = 390
    
            'The increment
            theIncrWidth = 2400 / theWidth
            theIncrHeight = 2400 / theHeight
    
            'Calculating new stage position
            xPosNew = (X * theIncrWidth) + xPosNow - 1200
            yPosNew = 1600 - ((Y * theIncrHeight) - yPosNow + 400)
            xPosNewStr = CStr(xPosNew)
            yPosNewStr = CStr(yPosNew)
    
    
            'set the new stage position
            PubMM.SetMMVariable "Device.Stage.XPosition", xPosNew
            PubMM.SetMMVariable "Device.Stage.YPosition", yPosNew
    
            'Adding to the text box
            Text4.Text = xPosNewStr
            Text6.Text = yPosNewStr
    
            'making a box
            Picture1.Line (X - 32.5, Y + 32.5)-(X + 32.5, Y + 32.5), vbRed
            Picture1.Line (X + 32.5, Y + 32.5)-(X + 32.5, Y - 32.5), vbRed
            Picture1.Line (X - 32.5, Y - 32.5)-(X + 32.5, Y - 32.5), vbRed
            Picture1.Line (X - 32.5, Y + 32.5)-(X - 32.5, Y - 32.5), vbRed
        
        End If
        
        '8x8 mosaic - navigation
        If Option6.Value = True Then
        
            'Width of image in pixel (picture box units)
            theWidth = 390
            theHeight = 390
    
            'The increment
            theIncrWidth = 3200 / theWidth
            theIncrHeight = 3200 / theHeight
    
            'Calculating new stage position
            xPosNew = (X * theIncrWidth) + xPosNow - 1600
            yPosNew = 2000 - ((Y * theIncrHeight) - yPosNow + 400)
            xPosNewStr = CStr(xPosNew)
            yPosNewStr = CStr(yPosNew)
    
    
            'set the new stage position
            PubMM.SetMMVariable "Device.Stage.XPosition", xPosNew
            PubMM.SetMMVariable "Device.Stage.YPosition", yPosNew
    
            'Adding to the text box
            Text4.Text = xPosNewStr
            Text6.Text = yPosNewStr
    
            'making a box
            Picture1.Line (X - 24.375, Y + 24.375)-(X + 24.375, Y + 24.375), vbRed
            Picture1.Line (X + 24.375, Y + 24.375)-(X + 24.375, Y - 24.375), vbRed
            Picture1.Line (X - 24.375, Y - 24.375)-(X + 24.375, Y - 24.375), vbRed
            Picture1.Line (X - 24.375, Y + 24.375)-(X - 24.375, Y - 24.375), vbRed
            
        End If
            
        '12x12 mosaic - navigation
        If Option7.Value = True Then
        
            'Width of image in pixel (picture box units)
            theWidth = 390
            theHeight = 390
    
            'The increment
            theIncrWidth = 4800 / theWidth
            theIncrHeight = 4800 / theHeight
    
            'Calculating new stage position
            xPosNew = (X * theIncrWidth) + xPosNow - 2400
            yPosNew = 2800 - ((Y * theIncrHeight) - yPosNow + 400)
            xPosNewStr = CStr(xPosNew)
            yPosNewStr = CStr(yPosNew)
    
    
            'set the new stage position
            PubMM.SetMMVariable "Device.Stage.XPosition", xPosNew
            PubMM.SetMMVariable "Device.Stage.YPosition", yPosNew
    
            'Adding to the text box
            Text4.Text = xPosNewStr
            Text6.Text = yPosNewStr
    
            'making a box
            Picture1.Line (X - 16.25, Y + 16.25)-(X + 16.25, Y + 16.25), vbRed
            Picture1.Line (X + 16.25, Y + 16.25)-(X + 16.25, Y - 16.25), vbRed
            Picture1.Line (X - 16.25, Y - 16.25)-(X + 16.25, Y - 16.25), vbRed
            Picture1.Line (X - 16.25, Y + 16.25)-(X - 16.25, Y - 16.25), vbRed
        
        
        End If
        
        '3x3 mosaic - navigation
        If Option10.Value = True Then
        
            'Width of image in pixel (picture box units)
            theWidth = 390
            theHeight = 390
    
            'The increment
            theIncrWidth = 1200 / theWidth
            theIncrHeight = 1200 / theHeight
    
            'Calculating new stage position
            xPosNew = (X * theIncrWidth) + xPosNow - 600
            yPosNew = 1000 - ((Y * theIncrHeight) - yPosNow + 400)
            xPosNewStr = CStr(xPosNew)
            yPosNewStr = CStr(yPosNew)
    
    
            'set the new stage position
            PubMM.SetMMVariable "Device.Stage.XPosition", xPosNew
            PubMM.SetMMVariable "Device.Stage.YPosition", yPosNew
    
            'Adding to the text box
            Text4.Text = xPosNewStr
            Text6.Text = yPosNewStr
    
            'making a box
            Picture1.Line (X - 65, Y + 65)-(X + 65, Y + 65), vbRed
            Picture1.Line (X + 65, Y + 65)-(X + 65, Y - 65), vbRed
            Picture1.Line (X - 65, Y - 65)-(X + 65, Y - 65), vbRed
            Picture1.Line (X - 65, Y + 65)-(X - 65, Y - 65), vbRed
        
        
        End If
        
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '''''''''''''''''''Case When You Are Drawing ROIs''''''''''''''''''''''''''''''
 
    ElseIf Option2.Value = True Then
    
    'getting the value of the button
    Dim theButtonMarker As Integer
    theButtonMarker = Button
   
    'The right click to make the mask

      If theButtonMarker = 2 Then
      
            'Some Definitions
            Dim theSlope As Double
            Dim theYinter As Double
            Dim xChange As Integer
            Dim yChange As Integer
            Dim theScaleNow As Integer
            Dim theMaskToBe As Long
            Dim xDrawDouble(100) As Double
            Dim yDrawDouble(100) As Double
            Dim idxStart As Integer
            Dim idxEnd As Integer
            Dim idxStartY As Integer
            Dim idxEndY As Integer
            Dim idxStartY2 As Integer
            Dim idxEndY2 As Integer
            Dim ActualWidth As Integer
            Dim ActualHeight As Integer
            Dim xEdge(10000) As Integer
            Dim yEdge(10000) As Integer
            Dim EdgeCounter As Integer
            
            'Draw all previous ROIs
            
            Text4.Text = CStr(ROINum)
            Text6.Text = CStr(MasterROICounter)
            
            If ROINum > 0 Then
            
                'putting a picture in gui
                Picture1.Picture = LoadPicture("D:\Users\JohnE\VisualBasicTests\TileTest2.bmp")

                'forcing the image to fit in the box
                Picture1.ScaleMode = 3
                Picture1.AutoRedraw = True
                Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight

                ' draw previous ROIs
                For u = 1 To ROINum
                    For s = 0 To (MasterROICounter + 1)
                            If IdxAllROIs(s + 1) = u And IdxAllROIs(s) = u Then
                               Picture1.Line (xDrawAllROIs(s), yDrawAllROIs(s))-(xDrawAllROIs(s + 1), yDrawAllROIs(s + 1)), vbRed
                            End If
                    Next s
                 Next u
    
                
            End If
            
            
            'Initializing the edge arrays
            For q = 1 To 10000
                xEdge(q - 1) = 0
                yEdge(q - 1) = 0
            Next q
            
            'Initializing Edge Counter
            EdgeCounter = 0
            
            'making the last mouse clicked position equal to the first to insure that polygon is closed
            xDraw(theCounter - 1) = xDraw(0)
            yDraw(theCounter - 1) = yDraw(0)
            
            're-drawing polygon to make sure that I have it - debugging
            For m = 1 To (theCounter - 1)
                  Picture1.Line (xDraw(m - 1), yDraw(m - 1))-(xDraw(m), yDraw(m)), vbGreen
            Next m
            
            'Loading the Big Image (the tiled mosaic)
            If Option3.Value = True Then ' 4x2 tiled image
                PubMM.CreateImage 2048, 1024, 16, "theMask", theMaskToBe
            ElseIf Option4.Value = True Then '4x4 tiled image
                PubMM.CreateImage 2048, 2048, 16, "theMask", theMaskToBe
            ElseIf Option5.Value = True Then '6x6 tiled image
                PubMM.CreateImage 3072, 3072, 16, "theMask", theMaskToBe
            ElseIf Option6.Value = True Then '8x8 tiled image
                PubMM.CreateImage 4096, 4096, 16, "theMask", theMaskToBe
            ElseIf Option7.Value = True Then '12x12 tiled image
                PubMM.CreateImage 6144, 6144, 16, "theMask", theMaskToBe
            ElseIf Option10.Value = True Then '3x3 tiled image
                PubMM.CreateImage 1536, 1536, 16, "theMask", theMaskToBe
            End If
                
            'Scaling
            For f = 0 To (theCounter - 1)
                If Option3.Value = True Then ' 4x2 tiled image
                    xDrawDouble(f) = CDbl(xDraw(f) * 3.88)
                    yDrawDouble(f) = CDbl(yDraw(f) * 3.9)
                ElseIf Option4.Value = True Then '4x4 tiled image
                    xDrawDouble(f) = CDbl(xDraw(f) * 5.25)
                    yDrawDouble(f) = CDbl(yDraw(f) * 5.25)
                ElseIf Option5.Value = True Then '6x6 tiled image
                    xDrawDouble(f) = CDbl(xDraw(f) * 7.88)
                    yDrawDouble(f) = CDbl(yDraw(f) * 7.88)
                ElseIf Option6.Value = True Then '8x8 tiled image
                    xDrawDouble(f) = CDbl(xDraw(f) * 10.5)
                    yDrawDouble(f) = CDbl(yDraw(f) * 10.5)
                ElseIf Option7.Value = True Then '12x12 tiled image
                    xDrawDouble(f) = CDbl(xDraw(f) * 15.75)
                    yDrawDouble(f) = CDbl(yDraw(f) * 15.75)
                ElseIf Option10.Value = True Then '3x3 tiled image
                    xDrawDouble(f) = CDbl(xDraw(f) * 3.94)
                    yDrawDouble(f) = CDbl(yDraw(f) * 3.94)
                End If
            Next f
            
            'adding a scale factor if user selects the TIRF dichroic
            For n = 0 To (theCounter - 1)
                If Option12.Value = True Then
                    xDrawDouble(n) = xDrawDouble(n) + 48
                    yDrawDouble(n) = yDrawDouble(n) + 94 ' was 166
                End If
            Next n
            
            'Drawing the outline of the mask
            For k = 1 To (theCounter - 1)
            
                    If CInt(xDrawDouble(k - 1)) <> CInt(xDrawDouble(k)) Then
                        'the Line
                        theSlope = (yDrawDouble(k - 1) - yDrawDouble(k)) / (xDrawDouble(k - 1) - xDrawDouble(k))
                        theYinter = yDrawDouble(k - 1) - (theSlope * xDrawDouble(k - 1))
                    End If
                    
                    If theSlope = 0 Then
                        theSlope = 1
                    End If
                        
                    
                    'indices
                    idxStart = CInt(xDrawDouble(k - 1))
                    idxEnd = CInt(xDrawDouble(k))
                    idxStartY = CInt(yDrawDouble(k - 1))
                    idxEndY = CInt(yDrawDouble(k))
                    
                    'drawing the outline of the mask
                    If idxStartY < idxEndY And idxStart <> idxEnd Then
                        For b = idxStartY To idxEndY
                            yChange = CInt(b)
                            xChange = CInt((b - theYinter) / theSlope)
                            'PubMM.WritePixel theMaskToBe, xChange, yChange, 1
                            
                            'Storing edge coordinates
                            xEdge(EdgeCounter) = xChange
                            yEdge(EdgeCounter) = yChange
                            EdgeCounter = EdgeCounter + 1
                            
                        Next b
                    ElseIf idxStartY > idxEndY And idxStart <> idxEnd Then
                        For c = idxEndY To idxStartY
                            yChange = CInt(c)
                            xChange = CInt((c - theYinter) / theSlope)
                           ' PubMM.WritePixel theMaskToBe, xChange, yChange, 1
                            
                            'Storing edge coordinates
                            xEdge(EdgeCounter) = xChange
                            yEdge(EdgeCounter) = yChange
                            EdgeCounter = EdgeCounter + 1
                            
                        Next c
                    ElseIf idxStartY = idxEndY And idxStart <> idxEnd Then
                            
                        If idxStart > idxEnd Then
                            For D = idxEnd To idxStart
                                yChange = CInt(idxStartY)
                                xChange = CInt(D)
                               ' PubMM.WritePixel theMaskToBe, xChange, yChange, 1
                                
                                'Storing edge coordinates
                                xEdge(EdgeCounter) = xChange
                                yEdge(EdgeCounter) = yChange
                                EdgeCounter = EdgeCounter + 1
                            
                            Next D
                        Else
                            For f = idxStart To idxEnd
                                yChange = CInt(idxStartY)
                                xChange = CInt(f)
                                'PubMM.WritePixel theMaskToBe, xChange, yChange, 1
                                
                                'Storing edge coordinates
                                xEdge(EdgeCounter) = xChange
                                yEdge(EdgeCounter) = yChange
                                EdgeCounter = EdgeCounter + 1
                                
                            Next f
                        End If
                                
                    End If
                    
                    
                    If idxStart > idxEnd And idxStartY <> idxEndY Then
                        For a = idxEnd To idxStart
                            xChange = CInt(a)
                            yChange = CInt((a * theSlope) + theYinter)
                           ' PubMM.WritePixel theMaskToBe, xChange, yChange, 1
                            
                            'Storing edge coordinates
                            xEdge(EdgeCounter) = xChange
                            yEdge(EdgeCounter) = yChange
                            EdgeCounter = EdgeCounter + 1
                            
                        Next a
                    ElseIf idxStart < idxEnd And idxStartY <> idxEndY Then
                       For a = idxStart To idxEnd
                            xChange = CInt(a)
                            yChange = CInt((a * theSlope) + theYinter)
                           ' PubMM.WritePixel theMaskToBe, xChange, yChange, 1
                            
                            'Storing edge coordinates
                            xEdge(EdgeCounter) = xChange
                            yEdge(EdgeCounter) = yChange
                            EdgeCounter = EdgeCounter + 1
                            
                        Next a
                    ElseIf idxStart = idxEnd And idxStartY <> idxEndY Then
                        If yDrawDouble(k - 1) > yDrawDouble(k) Then
                            idxStartY2 = yDrawDouble(k)
                            idxEndY2 = yDrawDouble(k - 1)
                        Else
                            idxStartY2 = yDrawDouble(k - 1)
                            idxEndY2 = yDrawDouble(k)
                        End If
                            For a = idxStartY2 To idxEndY2
                              xChange = CInt(xDrawDouble(k))
                              yChange = CInt(a)
                             ' PubMM.WritePixel theMaskToBe, xChange, yChange, 1
                              
                              'Storing edge coordinates
                              xEdge(EdgeCounter) = xChange
                              yEdge(EdgeCounter) = yChange
                              EdgeCounter = EdgeCounter + 1
                              
                            Next a
                    End If
                    
                
            Next k

            'ScanLine Filling
            Count1 = ScanLineFill(xEdge, yEdge, EdgeCounter)
            
            'resetting
            'initializing the counter in which I keep track of drawing of coordinates
            theCounter = 0
    
            'initializing the arrays that hold xy coordinates of drawing
            For i = 0 To 99
                xDraw(i) = 0
                yDraw(i) = 0
            Next i
            
            
        End If
    
    
        'record only left clicks
        If theButtonMarker = 2 Then Exit Sub
        
           'get xy coordinates
            XStart = X
            YStart = Y
            DrawMode = 7
            
            'store xy coordinates
            xDraw(theCounter) = X
            yDraw(theCounter) = Y
            
            'Draw a line as you go
            If theCounter > 0 Then
            
                Picture1.Line (xDraw(theCounter - 1), yDraw(theCounter - 1))-(xDraw(theCounter), yDraw(theCounter)), vbRed
            
            
            End If
            
            'iterate counter
            theCounter = theCounter + 1
            

    
    End If
    
    'adding some debugging output here
    'Dim xDebugKen As Double
    'Dim yDebugKen As Double
    'PubMM.GetMMVariable "Device.Stage.XPosition", xDebugKen
    'PubMM.GetMMVariable "Device.Stage.YPosition", yDebugKen
    'Text4.Text = CStr(xDebugKen)
    'Text6.Text = CStr(yDebugKen)
    
    
End Sub

Private Sub Form_Load()

   'This is some code to populate the combo box for channel 1 - for live and acquisition
    Combo1.AddItem "GFP 100%"
    Combo1.AddItem "DAPI 100%"
    Combo1.AddItem "Cy5 100%"
    Combo1.AddItem "TxRd 100%"
    
    'This is the code to populate the combo box for tiled imaging - Channel 1
    Combo2.AddItem "GFP 25%"
    Combo2.AddItem "GFP 50%"
    Combo2.AddItem "GFP 100%"
    Combo2.AddItem "TxRd 25%"
    Combo2.AddItem "TxRd 50%"
    Combo2.AddItem "TxRd 100%"
    
    'This is the code to populate the combo box for tiled imaging - Channel 2
    Combo3.AddItem "TxRd 25%"
    Combo3.AddItem "TxRd 50%"
    Combo3.AddItem "TxRd 100%"
    Combo3.AddItem "GFP 25%"
    Combo3.AddItem "GFP 50%"
    Combo3.AddItem "GFP 100%"
    

End Sub

Function FloodFill(ByVal xc As Integer, ByVal yc As Integer, x1 As Integer, y1 As Integer, ByVal New1 As Integer, ByVal Old1 As Integer, ByVal LowX As Integer, ByVal HighX As Integer, ByVal LowY As Integer, ByVal HighY As Integer) As Single

'initialization
Dim CurrIm As Long
Dim Pix1 As Integer
Dim theReturn As Single

'Get Current Image
PubMM.GetCurrentImage CurrIm

'keeping off of the edge
If x1 < LowX - 10 Or x1 > HighX + 10 Then
    x1 = xc
End If

If y1 < LowY - 10 Or y1 > HighY + 10 Then
    y1 = yc
End If

'Read Pixel Value
PubMM.ReadPixel CurrIm, CInt(x1), CInt(y1), Pix1

If Pix1 = Old1 Then
        
    'Write the new value
    PubMM.WritePixel CurrIm, CInt(x1), CInt(y1), New1
    
    'recursive calls
    theReturn = FloodFill(xc, yc, x1 + 1, y1, New1, Old1, LowX, HighX, LowY, HighY)
    theReturn = FloodFill(xc, yc, x1, y1 + 1, New1, Old1, LowX, HighX, LowY, HighY)
    theReturn = FloodFill(xc, yc, x1 - 1, y1, New1, Old1, LowX, HighX, LowY, HighY)
    theReturn = FloodFill(xc, yc, x1, y1 - 1, New1, Old1, LowX, HighX, LowY, HighY)
    


End If

FloodFill = 10

End Function

Function ScanLineFill(ByVal xEdgeSend1 As Variant, ByVal yEdgeSend1 As Variant, ByVal SizeEdge1 As Integer) As Single

'initialization
Dim CurrIm As Long
Dim Pix1 As Integer
Dim theCounter1 As Integer
Dim xKeep(1000) As Integer
Dim yKeep(1000) As Integer
Dim xKeepSort(1000) As Integer
Dim yKeepSort(1000) As Integer
Dim xKeepSortFinal(1000) As Integer
Dim yKeepSortFinal(1000) As Integer
Dim Temp As Integer
Dim Sorted As Boolean
Dim X As Integer
Dim xEdgeSend(10000) As Integer
Dim yEdgeSend(10000) As Integer
Dim SizeEdge As Integer
Dim LowY As Integer
Dim HighY As Integer

'Iterate ROI number when drawn
ROINum = ROINum + 1

'Get Current Image
PubMM.GetCurrentImage CurrIm

'Getting rid of duplicate enteries
For c = 0 To SizeEdge1 - 1
    For D = 0 To SizeEdge1 - 1
        If xEdgeSend1(c) = xEdgeSend1(D) And yEdgeSend1(c) = yEdgeSend1(D) And c <> D And xEdgeSend1(c) < 9000 Then
                xEdgeSend1(D) = 10000
                yEdgeSend1(D) = 10000
        End If
    Next D
Next c
SizeEdge = 0
For q = 0 To SizeEdge1 - 1
    If xEdgeSend1(q) < 9000 Then
        xEdgeSend(SizeEdge) = xEdgeSend1(q)
        yEdgeSend(SizeEdge) = yEdgeSend1(q)
        SizeEdge = SizeEdge + 1
    End If
Next q

'making the outline on the mask
For g = 0 To SizeEdge - 1
    PubMM.WritePixel CurrIm, xEdgeSend(g), yEdgeSend(g), 1
    PubMM.WritePixel CurrIm, xEdgeSend(g), yEdgeSend(g), 1
    PubMM.WritePixel CurrIm, xEdgeSend(g), yEdgeSend(g), 1
Next g

'get y extrema
HighY = 1
LowY = 3500
For t = 0 To (SizeEdge - 1)
            
    If yEdgeSend(t) > HighY Then
        HighY = yEdgeSend(t)
    End If
                
    If yEdgeSend(t) < LowY And yEdgeSend(t) > 0 Then
            LowY = yEdgeSend(t)
    End If
                
Next t

'Text4.Text = CStr(LowY)
'Text6.Text = CStr(HighY)

'Look along x dimension
For i = LowY To HighY

    'Initializing arrays to hold intersections
    For k = 0 To 999
        xKeep(k) = 0
        yKeep(k) = 0
        xKeepSort(k) = 0
        yKeepSort(k) = 0
        xKeepSortFinal(k) = 0
        yKeepSortFinal(k) = 0
    Next k

    'Initialize counter
    theCounter1 = 0

    'The intersections
    For j = 0 To SizeEdge - 1
        'Find the Intersections
        If yEdgeSend(j) = i Then
            xKeep(theCounter1) = xEdgeSend(j)
            theCounter1 = theCounter1 + 1
        End If
    Next j

    
    'Sorting
    For u = 1 To theCounter1 + 20
        For v = 0 To theCounter1 - 2
            If xKeep(v) > xKeep(v + 1) Then
             Temp = xKeep(v + 1)
              xKeep(v + 1) = xKeep(v)
              xKeep(v) = Temp
              Sorted = False
            End If
        Next v
    Next u
    
    'initializations
    Dim CurrX As Integer
    Dim CurrXAdd As Integer
    Dim RetCounter As Integer
    Dim RetCounterInit As Integer
    Dim NextCounter As Integer
    Dim TestFlag As Integer
    
    'initialize counter
    RetCounter = 0
    NextCounter = 0
    
    'Consolidating neighboring values
    For s = 0 To theCounter1 - 1
        
       If RetCounter <= (theCounter1 - 1) Then
       
            'current x
            CurrX = xKeep(RetCounter)
        
            'initial
            RetCounterInit = RetCounter
        
            'incrementing
            CurrXAdd = CurrX + 1
        
            For t = (RetCounter + 1) To (theCounter1 - 1)
                If CurrXAdd = xKeep(t) Then
            
                    'Iterate
                    CurrXAdd = CurrXAdd + 1
                
                    'Counter
                    RetCounter = RetCounter + 1
                End If
            Next t
            
            If RetCounter = RetCounterInit Then
                xKeepSort(NextCounter) = CurrX
                RetCounter = RetCounter + 1
            Else
                xKeepSort(NextCounter) = CurrXAdd
                RetCounter = RetCounter + 1
            End If
        
            'Iterate counter
            NextCounter = NextCounter + 1
        
        End If
        
    Next s
    
    'Masking
    For m = 1 To NextCounter
        If m = 2 Then
            For p = xKeepSort(m - 2) To xKeepSort(m - 1)
                'Write the new value
                PubMM.WritePixel CurrIm, CInt(p), CInt(i), 1
            Next p
        End If
        If m = 4 Then
            For p = xKeepSort(m - 2) To xKeepSort(m - 1)
                 'Write the new value
                 PubMM.WritePixel CurrIm, CInt(p), CInt(i), 1
            Next p
        End If
        If m = 6 Then
            For p = xKeepSort(m - 2) To xKeepSort(m - 1)
                'Write the new value
                PubMM.WritePixel CurrIm, CInt(p), CInt(i), 1
            Next p
        End If
        If m = 8 Then
            For p = xKeepSort(m - 2) To xKeepSort(m - 1)
                'Write the new value
                PubMM.WritePixel CurrIm, CInt(p), CInt(i), 1
            Next p
        End If
    Next m
 
 
Next i

'last step is median filter to clean up
PubMM.RunJournal "D:\Users\JohnE\VisualBasicTests\Journal Median Filter\JournalMedianFilter2.jnl"

'Get Current Image
PubMM.GetCurrentImage CurrIm

'save mask image as tif to make
If ROINum = 1 Then
    PubMM.SaveImage CurrIm, "D:\Users\JohnE\VisualBasicTests\ROIMask1.tif", False, 3
ElseIf ROINum = 2 Then
    PubMM.SaveImage CurrIm, "D:\Users\JohnE\VisualBasicTests\ROIMask2.tif", False, 3
ElseIf ROINum = 3 Then
    PubMM.SaveImage CurrIm, "D:\Users\JohnE\VisualBasicTests\ROIMask3.tif", False, 3
ElseIf ROINum = 4 Then
    PubMM.SaveImage CurrIm, "D:\Users\JohnE\VisualBasicTests\ROIMask4.tif", False, 3
ElseIf ROINum = 5 Then
    PubMM.SaveImage CurrIm, "D:\Users\JohnE\VisualBasicTests\ROIMask5.tif", False, 3
End If

'Close the big image
PubMM.CloseImage CurrIm

'Get Current Image
PubMM.GetCurrentImage CurrIm

'resave image as tif to make mask
PubMM.SaveImage CurrIm, "D:\Users\JohnE\VisualBasicTests\Tmp.tif", False, 3

'Close the big image
PubMM.CloseImage CurrIm

'Loading the final mask
'PubMM.LoadImage "D:\Users\JohnE\VisualBasicTests\TileTestMasked2.tif", 1

'Storing global information about the ROIs
For c = 0 To 99
    If xDraw(c) > 0 And yDraw(c) > 0 Then
        xDrawAllROIs(MasterROICounter) = xDraw(c)
        yDrawAllROIs(MasterROICounter) = yDraw(c)
        IdxAllROIs(MasterROICounter) = ROINum
        MasterROICounter = MasterROICounter + 1
    End If
Next c



'the return - visual basic egh.....
ScanLineFill = 10

End Function

Private Sub VScroll1_Change()

'setting the minimum and maximum
VScroll1.Min = 0
VScroll1.Max = 5

























End Sub

