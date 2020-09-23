VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Sphere covered with checkers. Written by Rick Ilyasov"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   FillColor       =   &H009DD1E6&
   Icon            =   "frmCheckerSphere.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Light Source"
      Height          =   1095
      Left            =   0
      TabIndex        =   40
      Top             =   5460
      Width           =   4935
      Begin VB.CheckBox chkRenderZY 
         Caption         =   "Render ZY"
         Height          =   255
         Left            =   1500
         TabIndex        =   55
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkRenderXZ 
         Caption         =   "Render XZ"
         Height          =   255
         Left            =   2640
         TabIndex        =   54
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtLSL 
         Height          =   285
         Left            =   3840
         TabIndex        =   48
         Text            =   "400"
         Top             =   360
         Width           =   555
      End
      Begin VB.TextBox txtLSZ 
         Height          =   285
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "3000"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtLSY 
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "3800"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtLSX 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "4000"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "LUM:"
         Height          =   195
         Left            =   3360
         TabIndex        =   47
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Z:"
         Height          =   255
         Left            =   2280
         TabIndex        =   45
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Y:"
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "X:"
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   420
         Width           =   495
      End
   End
   Begin TabDlg.SSTab tbTab 
      Height          =   5415
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9551
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      ForeColor       =   12582912
      TabCaption(0)   =   "Front XY"
      TabPicture(0)   =   "frmCheckerSphere.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picXY"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Side ZY"
      TabPicture(1)   =   "frmCheckerSphere.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picZY"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Top XZ"
      TabPicture(2)   =   "frmCheckerSphere.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picXZ"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picXZ 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         DrawWidth       =   2
         Height          =   4995
         Left            =   -74940
         ScaleHeight     =   4935
         ScaleWidth      =   4755
         TabIndex        =   39
         Top             =   60
         Width           =   4815
         Begin VB.PictureBox picLSXZ 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   3960
            Picture         =   "frmCheckerSphere.frx":0496
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   51
            Top             =   4140
            Width           =   240
         End
      End
      Begin VB.PictureBox picZY 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         DrawWidth       =   2
         Height          =   4995
         Left            =   -74940
         ScaleHeight     =   4935
         ScaleWidth      =   4755
         TabIndex        =   38
         Top             =   60
         Width           =   4815
         Begin VB.PictureBox picLSZY 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   4020
            Picture         =   "frmCheckerSphere.frx":07D8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   50
            Top             =   4260
            Width           =   240
         End
      End
      Begin VB.PictureBox picXY 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         DrawWidth       =   2
         Height          =   4995
         Left            =   60
         ScaleHeight     =   4935
         ScaleWidth      =   4755
         TabIndex        =   37
         Top             =   60
         Width           =   4815
         Begin VB.PictureBox picLSXY 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   360
            Picture         =   "frmCheckerSphere.frx":0B1A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   49
            Top             =   4320
            Width           =   240
         End
      End
   End
   Begin VB.TextBox txtCutHeight 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   7380
      TabIndex        =   33
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Y angle"
      ForeColor       =   &H8000000D&
      Height          =   855
      Index           =   1
      Left            =   6420
      TabIndex        =   20
      Top             =   1320
      Width           =   1455
      Begin VB.OptionButton optInnerYSin 
         Caption         =   "Sin"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optInnerYTan 
         Caption         =   "Tan"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optInnerYCos 
         Caption         =   "Cos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optInnerYAtn 
         Caption         =   "Atn"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "X angle"
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   4980
      TabIndex        =   25
      Top             =   1320
      Width           =   1455
      Begin VB.OptionButton optInnerXSin 
         Caption         =   "Sin"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optInnerXCos 
         Caption         =   "Cos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optInnerXTan 
         Caption         =   "Tan"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optInnerXAtn 
         Caption         =   "Atn"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.PictureBox picHighLight 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7260
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   4080
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5940
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   5940
      Width           =   1335
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Draw"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   5940
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Phase"
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   4980
      TabIndex        =   16
      Top             =   4560
      Width           =   2895
      Begin VB.TextBox txtPiRatio 
         Height          =   285
         Left            =   2280
         TabIndex        =   35
         Text            =   "0.5"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Pi ratio"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox picLowLight 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   7260
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   3480
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Y angle"
      ForeColor       =   &H8000000D&
      Height          =   855
      Index           =   0
      Left            =   6420
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
      Begin VB.OptionButton optYAtn 
         Caption         =   "Atn"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optYCos 
         Caption         =   "Cos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optYtan 
         Caption         =   "Tan"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optYSin 
         Caption         =   "Sin"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "X angle"
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   4980
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
      Begin VB.OptionButton optXAtn 
         Caption         =   "Atn"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optXTan 
         Caption         =   "Tan"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optXCos 
         Caption         =   "Cos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optXSin 
         Caption         =   "Sin"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtNumCells 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   5820
      TabIndex        =   1
      Text            =   "5"
      Top             =   720
      Width           =   495
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "Clear every time"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   5460
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808080&
      Caption         =   "Play around with these settings to achieve different shapes and effects"
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   4980
      TabIndex        =   53
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      Caption         =   "You may set the location of the light source in 3D space by clicking on the pictures in all 3 Views."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   52
      Top             =   6600
      Width           =   7875
   End
   Begin VB.Label Label6 
      Caption         =   "Cut height"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6540
      TabIndex        =   32
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Vertical Cut"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4980
      TabIndex        =   31
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Horizontal Cut"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4980
      TabIndex        =   30
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Light color                   (Double click to change)"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   4980
      TabIndex        =   14
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Dark color              (Double click to change)"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   4980
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Cell width"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4980
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Pixel3D
    X As Single
    Y As Single
    Z As Single
    R As Integer
    G As Integer
    B As Integer
End Type

Dim Stopped As Boolean
Dim HIGHLIGHT  As Long
Dim LOWLIGHT As Long

Dim LS As Pixel3D

Public Sub DrawSphere()

    Dim X1 As Single
    Dim X2 As Single
    
    Dim Y1 As Single
    Dim Y2 As Single
    
    Dim iTest As Integer
    
    Dim I As Single
    Dim pCCW As Single
    Dim pCW As Single
    
    Dim NumCells As Integer
    Dim CellStep As Single
    
    Dim sP As Double
    Dim sX As Single
    Dim sY As Single
    Dim StepCnt As Long
    
    Dim CutHeight As Integer
    Dim CutCenter As Single
    
    Dim Col As Long
    Dim NegCol As Long
    Dim Equals As Boolean
    Dim StepSumm As Double
    Dim ChangeCnt As Integer
    
    'Color Components
    Dim hR As Byte
    Dim hG As Byte
    Dim hB As Byte
    Dim lR As Byte
    Dim lG As Byte
    Dim lB As Byte
    
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    
    Dim iShift As Byte
    Dim iChannelOffset As Integer
    Dim iWait As Long
    Dim pHeight As Single
    Dim pWidth As Single
    Dim pWidthDiv2 As Single
    Dim pHeightDiv2 As Single
    
    Dim bSideView As Byte
    Dim iCutHeight As Integer
    
    Dim sngPiRatio As Single
    
    Dim binnerxcos As Byte
    Dim binnerxsin As Byte
    Dim binnerxtan As Byte
    Dim binnerxatn As Byte
    
    Dim binnerycos As Byte
    Dim binnerysin As Byte
    Dim binnerytan As Byte
    Dim binneryatn As Byte
    
    Dim bxcos As Byte
    Dim bxsin As Byte
    Dim bxtan As Byte
    Dim bxatn As Byte
    
    Dim bycos As Byte
    Dim bysin As Byte
    Dim bytan As Byte
    Dim byatn As Byte
    
    Dim PX As Pixel3D
    
    Dim sngDistance As Single
    Dim iAmountOfShadow As Integer
    Dim iLuminosity As Integer
    
    Dim bZY As Byte
    Dim bXZ As Byte
    
    bZY = chkRenderZY
    bXZ = chkRenderXZ
    
    sngPiRatio = txtPiRatio
    
    HIGHLIGHT = picHighLight.BackColor
    LOWLIGHT = picLowLight.BackColor
    
    hR = RofRGB(HIGHLIGHT)
    hG = GofRGB(HIGHLIGHT)
    hB = BofRGB(HIGHLIGHT)
    
    lR = RofRGB(LOWLIGHT)
    lG = GofRGB(LOWLIGHT)
    lB = BofRGB(LOWLIGHT)
        
    Col = HIGHLIGHT
    NegCol = LOWLIGHT
    
    NumCells = txtNumCells
    CellStep = 314 / NumCells
    
    pHeight = picXY.Height
    pWidth = picXY.Width
    pHeightDiv2 = pHeight / 2
    pWidthDiv2 = pWidth / 2
    
    'bSideView = optSideView
    iCutHeight = txtCutHeight
    
    binnerxcos = optInnerXCos
    binnerxsin = optInnerXSin
    binnerxtan = optInnerXTan
    binnerxatn = optInnerXAtn
    
    binnerycos = optInnerYCos
    binnerysin = optInnerYSin
    binnerytan = optInnerYTan
    binneryatn = optInnerYAtn
    
    bxcos = optXCos
    bxsin = optXSin
    bxtan = optXTan
    bxatn = optXAtn
    
    bycos = optYCos
    bysin = optYSin
    bytan = optYtan
    byatn = optYAtn
    
    'Set Light Source Position
    LS.X = txtLSX
    LS.Y = txtLSY
    LS.Z = txtLSZ  'The far point in the CUBE formed by the picture's dimensions
                'The depth of the cube from the Front side to the far side is equal to
                'The picture's width
                
    iLuminosity = txtLSL
    
    CutHeight = iCutHeight
    
    For I = 0 To (3.1415926 * sngPiRatio) Step 0.01
    
        If Stopped Then
            Exit For
        End If
        
        pCW = I
        pCCW = pCW * -1
        
        DoEvents
        
        If binnerxcos Then
            X1 = Cos(pCCW) * 2000 + pWidthDiv2
            X2 = Cos(pCW) * 2000 + pWidthDiv2
        ElseIf binnerxsin Then
            X1 = Sin(pCCW) * 2000 + pWidthDiv2
            X2 = Sin(pCW) * 2000 + pWidthDiv2
        ElseIf binnerxtan Then
            X1 = Tan(pCCW) * 2000 + pWidthDiv2
            X2 = Tan(pCW) * 2000 + pWidthDiv2
        ElseIf binnerxatn Then
            X1 = Atn(pCCW) * 2000 + pWidthDiv2
            X2 = Atn(pCW) * 2000 + pWidthDiv2
        End If

        If binnerycos Then
            Y1 = Cos(pCCW) * 2000 + pWidthDiv2
            Y2 = Cos(pCW) * 2000 + pWidthDiv2
        ElseIf binnerysin Then
            Y1 = Sin(pCCW) * 2000 + pWidthDiv2
            Y2 = Sin(pCW) * 2000 + pWidthDiv2
        ElseIf binnerytan Then
            Y1 = Tan(pCCW) * 2000 + pWidthDiv2
            Y2 = Tan(pCW) * 2000 + pWidthDiv2
        ElseIf binneryatn Then
            Y1 = Atn(pCCW) * 2000 + pWidthDiv2
            Y2 = Atn(pCW) * 2000 + pWidthDiv2
        End If
                
        If StepCnt Mod 5 = 0 Then
            If Col = LOWLIGHT Then
                Col = HIGHLIGHT
            Else
                Col = LOWLIGHT
            End If
        End If
            
        StepCnt = StepCnt + 1
        For sP = -3.14 / 2 To 3.14 / 2 Step 0.01
            
            CutCenter = Y1

            If bxcos Then
                sX = Cos(sP) * (Abs(X2 - X1) / 2) + pWidthDiv2
            ElseIf bxsin Then
                sX = Sin(sP) * (Abs(X2 - X1) / 2) + pWidthDiv2
            ElseIf bxtan Then
                sX = Tan(sP) * (Abs(X2 - X1) / 2) + pWidthDiv2
            ElseIf bxatn Then
                sX = Atn(sP) * (Abs(X2 - X1) / 2) + pWidthDiv2
            End If
            If bycos Then
                sY = Cos(sP) * (CutHeight) + CutCenter
            ElseIf bysin Then
                sY = Sin(sP) * (CutHeight) + CutCenter
            ElseIf bytan Then
                sY = Tan(sP) * (CutHeight) + CutCenter
            ElseIf byatn Then
                sY = Atn(sP) * (CutHeight) + CutCenter
            End If

            If StepSumm * 200 Mod NumCells = 0 Then
                If ChangeCnt Mod 2 = 0 Then
                    If NegCol = LOWLIGHT Then
                        NegCol = HIGHLIGHT
                        iShift = 1
                    Else
                        NegCol = LOWLIGHT
                        iShift = 0
                    End If
                Else
                    If Col = LOWLIGHT Then
                        NegCol = HIGHLIGHT
                        iShift = 1
                    Else
                        NegCol = LOWLIGHT
                        iShift = 0
                    End If
                End If
                ChangeCnt = ChangeCnt + 1
            End If

            PX.X = sX
            PX.Y = sY
            PX.Z = Sin(sP) * (Abs(X2 - X1) / 2) + pHeightDiv2
            
            Select Case iShift
                Case 0
                    R = lR: G = lG: B = lB
                Case 1
                    R = hR: G = hG: B = hB
            End Select
            
            sngDistance = Calculate3DDistance(PX, LS)
            
            iAmountOfShadow = 255 - (iLuminosity - (sngDistance / 10))

            R = R - iAmountOfShadow: If R < 0 Then R = 0: If R > 255 Then R = 255
            G = G - iAmountOfShadow: If G < 0 Then G = 0: If G > 255 Then G = 255
            B = B - iAmountOfShadow: If B < 0 Then B = 0: If B > 255 Then B = 255
                        
            picXY.PSet (PX.X, PX.Y), RGB(R, G, B)
            If bZY Then
                picZY.PSet (PX.Z, PX.Y), RGB(R, G, B)
            End If
            If bXZ Then
                picXZ.PSet (PX.X, PX.Z), RGB(R, G, B)
            End If
            
            StepSumm = StepSumm + 0.01

        Next

        StepCnt = StepCnt + 1
        For sP = 3.14 + 3.14 / 2 To 3.14 / 2 Step -0.01

            CutCenter = Y1

            If optXCos Then
                sX = Cos(sP) * (Abs(X2 - X1) / 2) + pWidthDiv2
            ElseIf optXSin Then
                sX = Sin(sP) * (Abs(X2 - X1) / 2) + pWidthDiv2
            ElseIf optXTan Then
                sX = Tan(sP) * (Abs(X2 - X1) / 2) + pWidthDiv2
            ElseIf optXAtn Then
                sX = Atn(sP) * (Abs(X2 - X1) / 2) + pWidthDiv2
            End If
            If optYCos Then
                sY = Cos(sP) * (CutHeight) + CutCenter
            ElseIf optYSin Then
                sY = Sin(sP) * (CutHeight) + CutCenter
            ElseIf optYtan Then
                sY = Tan(sP) * (CutHeight) + CutCenter
            ElseIf optYAtn Then
                sY = Atn(sP) * (CutHeight) + CutCenter
            End If

            If StepSumm * 200 Mod NumCells = 0 Then
                If ChangeCnt Mod 2 = 0 Then
                    If NegCol = LOWLIGHT Then
                        NegCol = HIGHLIGHT
                        iShift = 1
                    Else
                        NegCol = LOWLIGHT
                        iShift = 0
                    End If
                Else
                    If Col = LOWLIGHT Then
                        NegCol = HIGHLIGHT
                        iShift = 1
                    Else
                        NegCol = LOWLIGHT
                        iShift = 0
                    End If
                End If
                ChangeCnt = ChangeCnt + 1
            End If
            
            PX.X = sX
            PX.Y = sY
            PX.Z = Sin(sP) * (Abs(X2 - X1) / 2) + pHeightDiv2
            
            Select Case iShift
                Case 0
                    R = lR: G = lG: B = lB
                Case 1
                    R = hR: G = hG: B = hB
            End Select

            sngDistance = Calculate3DDistance(PX, LS)
            
            iAmountOfShadow = 255 - (iLuminosity - (sngDistance / 10))

            R = R - iAmountOfShadow: If R < 0 Then R = 0: If R > 255 Then R = 255
            G = G - iAmountOfShadow: If G < 0 Then G = 0: If G > 255 Then G = 255
            B = B - iAmountOfShadow: If B < 0 Then B = 0: If B > 255 Then B = 255

            picXY.PSet (PX.X, PX.Y), RGB(R, G, B)
            If bZY Then
                picZY.PSet (PX.Z, PX.Y), RGB(R, G, B)
            End If
            If bXZ Then
                picXZ.PSet (PX.X, PX.Z), RGB(R, G, B)
            End If
            
            StepSumm = StepSumm + 0.01

        Next

        ChangeCnt = 0
        StepSumm = 0
    Next
    
    Stopped = False
    
End Sub

Private Function Calculate3DDistance(PT1 As Pixel3D, PT2 As Pixel3D) As Single

    Dim XYhypothenuse As Single
    Dim ZYhypothenuse As Single
    Dim ZXhypothenuse As Single
    
    Dim Xside As Single
    Dim Yside As Single
    Dim Zside As Single
    
    XYhypothenuse = Sqr(Abs(PT2.X - PT1.X) ^ 2 + Abs(PT2.Y - PT1.Y) ^ 2)
    ZYhypothenuse = Sqr(Abs(PT2.Z - PT1.Z) ^ 2 + Abs(PT2.Y - PT1.Y) ^ 2)
    ZXhypothenuse = Sqr(Abs(PT2.Z - PT1.Z) ^ 2 + Abs(PT2.X - PT1.X) ^ 2)
    
    Xside = Abs(PT2.X - PT1.X)
    Yside = Abs(PT2.Y - PT1.Y)
    Zside = Abs(PT2.Z - PT1.Z)
    
    Calculate3DDistance = XYhypothenuse + (ZYhypothenuse - Yside)
    
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdToggle_Click()
    
    If cmdToggle.Caption = "Draw" Then
        
        cmdToggle.Caption = "Stop"
        If chkClear Then
            picXY.Cls
            picZY.Cls
            picXZ.Cls
        End If
        
        DrawSphere
        
        cmdToggle.Caption = "Draw"
        
    Else
        Stopped = True
        cmdToggle.Caption = "Draw"
    End If
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    optInnerYCos = True
    optInnerXSin = True
    optXCos = True
    optYSin = True
    
    txtNumCells = 5
    txtCutHeight = 0
    txtPiRatio = 1
    chkClear = 1
    
    txtLSX = picLSXY.Left + picLSXY.Width / 2
    txtLSY = picLSXY.Top + picLSXY.Height / 2
    txtLSZ = picLSZY.Left + picLSZY.Width / 2
    picLSZY.Top = txtLSY - picLSZY.Height / 2
    picLSXZ.Left = txtLSX - picLSXZ.Width / 2
    picLSXZ.Top = txtLSZ - picLSXZ.Height / 2
    
End Sub

Private Sub picHighLight_Click()
    CommonDialog1.ShowColor
    picHighLight.BackColor = CommonDialog1.Color
End Sub

Private Sub picLowLight_Click()
    CommonDialog1.ShowColor
    picLowLight.BackColor = CommonDialog1.Color
End Sub


Public Function RofRGB(RGBCol As Long) As Long
    RofRGB = (RGBCol Mod 65536) Mod 256
End Function

Public Function GofRGB(RGBCol As Long) As Long
    
    Dim RG As Long
    Dim R As Integer
    
    RG = RGBCol Mod 65536
    R = RG Mod 256
    GofRGB = (RG - R) / 256

End Function

Public Function BofRGB(RGBCol As Long) As Long
    BofRGB = (RGBCol - (RGBCol Mod 65536)) / 65536
End Function

Private Sub picXY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picLSXY.Left = X - picLSXZ.Width / 2
    picLSXY.Top = Y - picLSXZ.Height / 2

    picLSXZ.Left = picLSXY.Left
    picLSZY.Top = picLSXY.Top
    
    txtLSX = X
    txtLSY = Y

End Sub

Private Sub picZY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picLSZY.Left = X - picLSZY.Width / 2
    picLSZY.Top = Y - picLSZY.Height / 2
    
    picLSXY.Top = picLSZY.Top
    picLSXZ.Top = picLSZY.Top
        
    txtLSZ = picZY.ScaleWidth - X
    txtLSY = Y

End Sub

Private Sub picXZ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picLSXZ.Left = X - picLSXZ.Width / 2
    picLSXZ.Top = Y - picLSXZ.Height / 2

    picLSXY.Left = picLSXZ.Left
    picLSZY.Left = picXZ.Height - (picLSXZ.Top + picLSXZ.Height / 2) - picLSZY.Width / 2
    
    txtLSX = X
    txtLSZ = Y

End Sub
