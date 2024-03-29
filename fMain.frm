VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "3D smoothed particles hydrodynamics"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   682
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1008
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "R"
      Height          =   615
      Left            =   13200
      TabIndex        =   19
      Top             =   6240
      Width           =   615
   End
   Begin VB.CheckBox chkRot 
      Caption         =   "Rotate CAM"
      Height          =   495
      Left            =   12360
      TabIndex        =   18
      Top             =   9720
      Width           =   1695
   End
   Begin VB.CommandButton cmdErase 
      Caption         =   "Clean 2"
      Height          =   375
      Index           =   1
      Left            =   13680
      TabIndex        =   17
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdErase 
      Caption         =   "Clean 1"
      Height          =   375
      Index           =   0
      Left            =   12360
      TabIndex        =   16
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CheckBox chkCOMG 
      Caption         =   "COM Gravity"
      Height          =   495
      Left            =   12360
      TabIndex        =   15
      ToolTipText     =   "Center of Mass gravity"
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CheckBox chkFaucet 
      Caption         =   "Faucet 2"
      Height          =   495
      Index           =   1
      Left            =   13680
      TabIndex        =   14
      Top             =   8040
      Width           =   1695
   End
   Begin VB.PictureBox picGravity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   12360
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   88
      TabIndex        =   12
      ToolTipText     =   "Right click for 0,0,0"
      Top             =   5520
      Width           =   1350
      Begin VB.Line Line1 
         X1              =   72
         X2              =   96
         Y1              =   72
         Y2              =   112
      End
   End
   Begin VB.CheckBox chkFaucet 
      Caption         =   "Faucet 1"
      Height          =   495
      Index           =   0
      Left            =   12360
      TabIndex        =   11
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CheckBox chkRG 
      Caption         =   "Rnd Gravity"
      Height          =   495
      Left            =   12360
      TabIndex        =   10
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CheckBox chkJPG 
      Caption         =   "Save PNG Frames"
      Height          =   495
      Left            =   12360
      TabIndex        =   9
      Top             =   9240
      Width           =   1695
   End
   Begin VB.TextBox txtMaxD 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   12360
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtNP 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   12360
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.HScrollBar ScrollDRAW 
      Height          =   255
      Left            =   12480
      Max             =   10
      TabIndex        =   3
      Top             =   2160
      Value           =   2
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11640
      Top             =   240
   End
   Begin VB.CheckBox chkDRAWP 
      Caption         =   "Draw Pairs"
      Height          =   495
      Left            =   12480
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(re)Start"
      Height          =   1335
      Left            =   12480
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   120
      ScaleHeight     =   345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   529
      TabIndex        =   0
      ToolTipText     =   "Clcik and Move to Rotate Camera (Right click to Reset)"
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label3 
      Caption         =   "Gravity: Click Pic."
      Height          =   375
      Left            =   12360
      TabIndex        =   13
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Max Dist:"
      Height          =   375
      Left            =   12360
      TabIndex        =   6
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "N Points:"
      Height          =   375
      Left            =   12360
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lSkip 
      Alignment       =   2  'Center
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SRF       As cCairoSurface
Private CC        As cCairoContext


Private RecomputeBOX As Boolean
Private InvGravScale As Single


Private Sub chkCOMG_Click()
    COMGravity = chkCOMG.Value = vbChecked
End Sub

Private Sub chkDRAWP_Click()
    DrawPairs = chkDRAWP.Value = vbChecked
End Sub




Private Sub chkFaucet_Click(Index As Integer)
    If Index = 0 Then DoFaucet1 = chkFaucet(Index).Value = vbChecked
    If Index = 1 Then DoFaucet2 = chkFaucet(Index).Value = vbChecked

End Sub

Private Sub chkJPG_Click()
    DoSaveFrames = chkJPG = vbChecked
End Sub

Private Sub chkRG_Click()
    rndGravity = chkRG.Value = vbChecked
End Sub

Private Sub chkRot_Click()
    CamRot = chkRot.Value = vbChecked
End Sub

Private Sub cmdErase_Click(Index As Integer)
    Dim I&, J&
    Dim tpx()     As Single
    Dim tpy()     As Single
    Dim tpz()     As Single

    Dim tVx()     As Single
    Dim tVy()     As Single
    Dim tVz()     As Single

    Dim tPhase()  As Long

    ReDim tpx(NP)
    ReDim tpy(NP)
    ReDim tpz(NP)
    ReDim tVx(NP)
    ReDim tVy(NP)
    ReDim tVz(NP)
    ReDim tPhase(NP)

    Index = Index + 1

    For I = 1 To NP
        If Phase(I) <> Index Then
            J = J + 1
            tpx(J) = px(I)
            tpy(J) = py(I)
            tpz(J) = pz(I)
            tVx(J) = vX(I)
            tVy(J) = vY(I)
            tVz(J) = vZ(I)
            tPhase(J) = Phase(I)
        End If
    Next

    px() = tpx()
    py() = tpy()
    pz() = tpz()
    vX() = tVx()
    vY() = tVy()
    vZ() = tVz()
    Phase() = tPhase()

    NP = J

    ReDim Preserve px(NP)
    ReDim Preserve py(NP)
    ReDim Preserve pz(NP)

    ReDim Preserve vX(NP)
    ReDim Preserve vY(NP)
    ReDim Preserve vZ(NP)

    ReDim Preserve pvX(NP)
    ReDim Preserve pvY(NP)
    ReDim Preserve pvZ(NP)

    ReDim Preserve Phase(NP)

End Sub

Private Sub Command2_Click()
    gTOX = (Rnd * 2 - 1) * GravScale
    gTOY = (Rnd * 2 - 1) * GravScale
    gTOZ = (Rnd * 2 - 1) * GravScale * 1.1
    If Rnd > 0.2 Then              '0.1
        If Abs(gTOX) > Abs(gTOY) And Abs(gTOX) > Abs(gTOZ) Then gTOY = 0: gTOZ = 0
        If Abs(gTOY) > Abs(gTOX) And Abs(gTOY) > Abs(gTOX) Then gTOX = 0: gTOZ = 0
        If Abs(gTOZ) > Abs(gTOY) And Abs(gTOZ) > Abs(gTOX) Then gTOY = 0: gTOX = 0

        If Rnd < 0.1 Then gTOX = 0: gTOY = 0: gTOZ = 0
    End If

    Line1.X1 = picGravity.Width * 0.5
    Line1.Y1 = picGravity.Height * 0.5
    Line1.X2 = Line1.X1 + gTOX * picGravity.Width * 0.5 * InvGravScale
    Line1.Y2 = Line1.Y1 + gTOY * picGravity.Height * 0.5 * InvGravScale

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Randomize Timer

    PIC.Cls

    Me.Caption = Me.Caption & " V." & App.Major


    If Dir(App.Path & "\Frames", vbDirectory) = vbNullString Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.*", vbArchive) <> vbNullString Then Kill App.Path & "\Frames\*.*"

    HH = 520
    WW = 4 / 3 * HH
    '    WW = 16 / 9 * HH


    If WW < HH Then ZZ = WW Else: ZZ = HH
    ZZ = ZZ * 0.5
    invZZ = 1 / ZZ

    WW = WW - (WW Mod 4)
    PIC.Height = HH
    PIC.Width = WW



    PIChDC = PIC.hDC

    ScrollDRAW.Value = 2
    txtNP.Text = 5000              '2000   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    txtMaxD.Text = 14              '25



    Dim ctr       As Control
    For Each ctr In Me
        If ctr.Name <> "PIC" And ctr.Name <> "Line1" And ctr.Name <> "Timer1" _
           And ctr.Name <> "Command2" Then
            ctr.Left = PIC.Width + 50
        End If
    Next

    chkFaucet(1).Left = chkFaucet(0).Left + chkFaucet(0).Width + 1
    cmdErase(1).Left = cmdErase(0).Left + cmdErase(0).Width + 10



    chkRG.Value = vbChecked
    rndGravity = True


    '-------------------------------------------------- CAMERA V3
    '    Scree.Size = Vec3(WW * 1, HH * 1, ZZ * 1)
    '    Scree.InvSize.X = 1 / Scree.Size.X
    '    Scree.InvSize.Y = 1 / Scree.Size.Y
    '    Scree.InvSize.Z = 1 / Scree.Size.Z
    '    Scree.Center = Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5)
    '    InitCamera Vec3(WW * 0.5, HH * 0.5, HH * 1.252), Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5)

    ' Camera V4
    '    CameraInit Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5 + WW * 0.7), _
         Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5), _
         Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5), Vec3(0, 1, 0)


    'V6
    Set CAMERA = New c3DEasyCam

    CAMERA.Init Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5 - WW * 0.7), _
                Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5), _
                Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5), Vec3(0, -1, 0)


End Sub



Private Sub Command1_Click()

    RecomputeBOX = True

    Set SRF = Cairo.CreateSurface(WW + 1, HH + 1, ImageSurface)
    Set CC = SRF.CreateContext
    CC.AntiAlias = CAIRO_ANTIALIAS_FAST

    '    Set SpatialGRID = New cSpatialGrid3D
    '    '    Set HASH3D = New cSpatialHash3D
    Set OCTTREE = New cOctTree


    H = Val(txtMaxD.Text)
    invH = 1 / H
    InvH2 = 1 / (H * H)
    MainLoop

End Sub


Public Sub MainLoop()

    Dim I         As Long
    Dim J         As Long


    Dim V         As Single
    Dim OnP       As String

    Dim X         As Single
    Dim Y         As Single
    Dim Z         As Single
    Dim InvZ      As Single


    Dim DrawOrderIDX() As Long
    Dim DistFromCamera() As Single

    ReDim DrawOrderIDX(1)
    ReDim DistFromCamera(1)

    Dim L         As Long
    Dim tV        As tVec3


    ' V5 BOX
    Dim LP1(1 To 12) As tVec3
    Dim LP2(1 To 12) As tVec3
    Dim Vis(1 To 12) As Boolean



    Dim DrawR     As Single

    Dim yyy       As Single
    Dim ppp       As Single




    Dim C         As cHashD


    Dim CamPosX   As Single
    Dim CamPosY   As Single
    Dim CamPosZ   As Single
    Dim DX As Single, DY As Single, DZ As Single

    Dim MAXNPairs As Long


    DrawR = H * 0.12

    '''OLDPointToScreen    If DrawR < 1.5 Then DrawR = 1.5


    DoLOOP = True

    NP = Val(txtNP.Text)


    SPH_InitConst


    'SpatialGRID.Init WW, HH, ZZ, h * 1

    OCTTREE.Setup -H, -H, -H, WW + H, HH + H, ZZ + H, 65    ', H '60    ' 50 ' 50 '60 '10


    '    HASH3D.constructor h * 1, NP
    '    Set C = New_c.HashD(NP)

    'GravScale = (h / 60) * invDT
    GravScale = (H / 70) * invDT   '2022

    InvGravScale = 1 / GravScale

    ReDim px(NP)
    ReDim py(NP)
    ReDim pz(NP)

    ReDim vX(NP)
    ReDim vY(NP)
    ReDim vZ(NP)

    ReDim pvX(NP)
    ReDim pvY(NP)
    ReDim pvZ(NP)


    gX = 0
    gY = 1
    gZ = 0



    For I = 1 To NP
        px(I) = Rnd * WW * 1
        py(I) = Rnd * HH * 1
        pz(I) = Rnd * ZZ * 1

        vX(I) = (Rnd * 2 - 1) * 1
        vY(I) = (Rnd * 2 - 1) * 1
        vZ(I) = (Rnd * 2 - 1) * 1

        '        X = PX(I) - WW * 0.5
        '        Z = PZ(I) - ZZ * 0.5
        '        A = Atan2(X, Z)
        '        D = Sqr(X * X + Z * Z)
        '
        '        vX(I) = Cos(A) * 0.05 * invDT * D
        '        vY(I) = 0
        '        vZ(I) = Sin(A) * 0.05 * invDT * D


        Phase(I) = 1

    Next

    Do

        '    CAMERA.Follow Vec3(pX(1), pY(1), pZ(1)), 0.03, 0.0125, 30000, 18000: RecomputeBOX = True
        '    CAMERA.Follow Vec3(COMx, COMy, COMz), 0.025, 0.0125, 80600, 40000: RecomputeBOX = True

        SPH_MOVE

        '------------------------
        '------------------------
        '------------------------MODO grid




        'SpatialGRID.ResetPoints
        ''------------------------
        '       SpatialGRID.InsertALLpoints px, py, pz
        ''------------------------


        OCTTREE.InsertALLpoints px, py, pz



        Dim tdt   As Single
        tdt = DT * 0.5
        ''''        ''        '-------using pvX '? Dampig?
        ''        For I = 1 To NP
        ''            pvX(I) = pX(I) + vX(I) * tdt
        ''            pvY(I) = pY(I) + vY(I) * tdt
        ''            pvZ(I) = pZ(I) + vZ(I) * tdt
        ''        Next
        ''        SpatialGRID.InsertALLpoints pvX, pvY, pvZ
        ''        '------------------------
        '------------------------


        'SpatialGRID.GetPairsWDist4 P1, P2, arrDX, arrDY, arrDZ, arrD, RetNofPairs
        ''        SpatialGRID.GetPairsWDist2 P1, P2, arrDX, arrDY, arrDZ, arrD, RetNofPairs



        OCTTREE.GetPairsWDist H, P1, P2, arrDX, arrDY, arrDZ, arrD, RetNofPairs, MAXNPairs



        '------------------------
        '------------------------
        '------------------------MODO HASH
        '''        Dim nr&
        '''
        '''        Dim DX as single, DY as single, DZ as single, D as single
        '''
        '''        HASH3D.InsertPoints2 pX, pY, pZ
        '''        RetNofPairs = 0
        '''        C.ReInit MaxNofPairs
        '''
        '''        For I = 1 To NP
        '''
        '''            HASH3D.Query Vec3(pX(I), pY(I), pZ(I)), h
        '''            For nr = 1 To HASH3D.querySize
        '''
        '''                J = HASH3D.QueryIds(nr)
        '''
        '''                If Not (C.Exists(J & I)) Then
        '''
        '''
        '''                    DX = pX(J) - pX(I)
        '''                    DY = pY(J) - pY(I)
        '''                    DZ = pZ(J) - pZ(I)
        '''                    D = DX * DX + DY * DY + DZ * DZ
        '''                    If D > 0 And D < (h * h) Then
        '''
        '''                        If Not (C.Exists(I & J)) Then
        '''                            C.Add I & J, ""
        '''
        '''                            RetNofPairs = RetNofPairs + 1
        '''
        '''                            If RetNofPairs > MaxNofPairs Then
        '''                                MaxNofPairs = RetNofPairs * 1.7
        '''                                ReDim Preserve P1(MaxNofPairs): ReDim Preserve P2(MaxNofPairs)
        '''                                ReDim Preserve arrDX(MaxNofPairs): ReDim Preserve arrDY(MaxNofPairs): ReDim Preserve arrDZ(MaxNofPairs)
        '''                                ReDim Preserve arrD(MaxNofPairs)
        '''                            End If
        '''
        '''                            P1(RetNofPairs) = I
        '''                            P2(RetNofPairs) = J
        '''                            arrDX(RetNofPairs) = DX
        '''                            arrDY(RetNofPairs) = DY
        '''                            arrDZ(RetNofPairs) = DZ
        '''                            arrD(RetNofPairs) = D
        '''                            If I = 0 And J = 0 Then Stop
        '''                        End If
        '''                    End If
        '''
        '''                End If
        '''
        '''
        '''            Next
        '''
        '''        Next
        '------------------------
        '------------------------
        '------------------------

        SPH_ComputePAIRS






        '----------------------

        If (CNT And RenderEvery) = 0& Then
            With CC


                .SetSourceColor 0
                .Paint

                ' DRAW BOX------------------------------ V5
                If RecomputeBOX Then

                    CAMERA.LineToScreen Vec3(0, 0, 0), Vec3(WW * 1, 0, 0), LP1(1), LP2(1), Vis(1)
                    CAMERA.LineToScreen Vec3(WW * 1, 0, 0), Vec3(WW * 1, HH * 1, 0), LP1(2), LP2(2), Vis(2)
                    CAMERA.LineToScreen Vec3(WW * 1, HH * 1, 0), Vec3(0, HH * 1, 0), LP1(3), LP2(3), Vis(3)
                    CAMERA.LineToScreen Vec3(0, HH * 1, 0), Vec3(0, 0, 0), LP1(4), LP2(4), Vis(4)

                    CAMERA.LineToScreen Vec3(0, 0, ZZ * 1), Vec3(WW * 1, 0, ZZ * 1), LP1(5), LP2(5), Vis(5)
                    CAMERA.LineToScreen Vec3(WW * 1, 0, ZZ * 1), Vec3(WW * 1, HH * 1, ZZ * 1), LP1(6), LP2(6), Vis(6)
                    CAMERA.LineToScreen Vec3(WW * 1, HH * 1, ZZ * 1), Vec3(0, HH * 1, ZZ * 1), LP1(7), LP2(7), Vis(7)
                    CAMERA.LineToScreen Vec3(0, HH * 1, ZZ * 1), Vec3(0, 0, ZZ * 1), LP1(8), LP2(8), Vis(8)


                    CAMERA.LineToScreen Vec3(0, 0, 0), Vec3(0, 0, ZZ * 1), LP1(9), LP2(9), Vis(9)
                    CAMERA.LineToScreen Vec3(WW * 1, 0, 0), Vec3(WW * 1, 0, ZZ * 1), LP1(10), LP2(10), Vis(10)
                    CAMERA.LineToScreen Vec3(WW * 1, HH * 1, 0), Vec3(WW * 1, HH * 1, ZZ * 1), LP1(11), LP2(11), Vis(11)
                    CAMERA.LineToScreen Vec3(0, HH * 1, 0), Vec3(0, HH * 1, ZZ * 1), LP1(12), LP2(12), Vis(12)
                    RecomputeBOX = False
                End If
                .SetSourceRGBA 0.5, 0.85, 0.5, 0.5    ' 0.35

                For L = 1 To 12
                    ' If LP1(L).Z > 0 And LP2(L).Z > 0 Then .MoveTo LP1(L).X, LP1(L).Y: .LineTo LP2(L).X, LP2(L).Y: .Stroke
                    If Vis(L) Then .MoveTo LP1(L).X, LP1(L).Y: .LineTo LP2(L).X, LP2(L).Y: .Stroke
                Next
                ' END DRAW BOX--------------------------------------------


                '--------------------------------------------------------
                '--------- DRAW Points ----------------------------------
                '--------------------------------------------------------


                ' DRAW ORDER -------
                If NP <> UBound(DrawOrderIDX) Then
                    ReDim DrawOrderIDX(NP)
                    For I = 1 To NP
                        DrawOrderIDX(I) = I
                    Next
                End If
                If NP > UBound(DistFromCamera) Then ReDim DistFromCamera(NP)


                With CAMERA.Position
                    CamPosX = .X
                    CamPosY = .Y
                    CamPosZ = .Z
                End With
                For I = 1 To NP
                    '                    '                    DrawOrderIDX(I) = I
                    '                    '                    'OK ! Project camera to point vector to CamNormFrontDIR (Front camera Vector)
                    '                    '                    'DistFromCamera(I) = dot3(DIFF3(Camera.position, Vec3(pX(I), pY(I), pZ(I))), CamNormFrontDIR) 'Camera V3
                    '                    '                    'Camera V4
                    '                    '                    'DistFromCamera(I) = DOT3(DIFF3(CAMERA.Position, Vec3(pX(I), pY(I), pZ(I))), CAMERA.Direction)
                    '                    'With DIFF3(CAMERA.Position, Vec3(pX(I), pY(I), pZ(I)))
                    '                    '     DistFromCamera(I) = -(.x * .x + .y * .y + .z * .z)
                    '                    'End With
                    '
                    '                    DX = pX(I) - CamPosX
                    '                    DY = pY(I) - CamPosY
                    '                    DZ = pZ(I) - CamPosZ
                    '                    DistFromCamera(I) = -(DX * DX + DY * DY + DZ * DZ)

                    '---- Speed up Sort:
                    J = DrawOrderIDX(I)
                    DX = px(J) - CamPosX
                    DY = py(J) - CamPosY
                    DZ = pz(J) - CamPosZ
                    DistFromCamera(I) = -(DX * DX + DY * DY + DZ * DZ)

                Next

                SORTSWAPS = 0
                'QuickSortSingle2 DistFromCamera(), DrawOrderIDX(), 1, NP
                QuickSortSingle3 DistFromCamera(), DrawOrderIDX(), 1, NP

                ' END DRAW ORDER -------

                For I = NP To 1 Step -1

                    J = DrawOrderIDX(I)
                    '
                    '                    'V = Pressure(J) * 0.075
                    '                    'V = Pressure(J) * 0.15    'V5
                    V = Pressure(J) * 0.12    '2022

                    'V = (vX(J) * vX(J) + vY(J) * vY(J) + vZ(J) * vZ(J)) * 0.0015




                    CAMERA.PointToScreenCoords px(J), py(J), pz(J), X, Y, Z, InvZ
                    'If Z > 0 Then
                    'If CAMERA.IsPointVisibleGap(Vec3(X, y, Z), 20) Then
                    If CAMERA.IsPointVisibleGap2(X, Y, Z, 20) Then




                        V = 0.4 + V - 0.0009 * Z    'Darker distance


                        ' V = V + Z * 50 - 0.2
                        If Phase(J) = 1& Then
                            'cyan
                            '.SetSourceRGBA 0.1 + V, 0.65 + V, 0.75 + V, 0.7  (V 1 , 2 )
                            '                      .SetSourceRGBA 0.02 + V, 0.5 + V, 0.6 + V, 0.85   '(V 3 )

                            '.SetSourceRGBA 0.01 + V, 0.45 + V, 0.55 + V, 0.8    '0.7  '(V 4)

                            '.SetSourceRGBA 0.015 + V, 0.5 + V, 0.6 + V, 0.8  '0.7  '(V 5)

                            .SetSourceRGBA 0.015 + V, 0.5 + V, 0.6 + V, 0.7    '0.55
                            '                            .SetSourceRGB 0.015 + V, 0.5 + V, 0.6 + V

                        Else
                            '.SetSourceRGBA 0.2 + V, 0.7 + V, 0.2 + V, 0.7

                            ' .SetSourceRGBA 0.15 + V, 0.6 + V, 0.15 + V, 0.8
                            .SetSourceRGBA 0.1 + V, 0.5 + V, 0.1 + V, 0.7    ' 0.55
                            ' .SetSourceRGB 0.1 + V, 0.5 + V, 0.1 + V

                        End If

                        '.Arc X, Y, 0.7 + DrawR * Z * 340   '(V3)
                        '.Arc X, Y, 0.7 + DrawR * InvZ * 450   'V4
                        .Arc X, Y, 1 + DrawR * InvZ * 400
                        .Fill
                    End If

                Next
                '--------------------------------------------------------
                '--------------------------------------------------------
                '--------------------------------------------------------


                '''                If DrawPairs Then
                '''                    .SetSourceRGBA 1, 1, 0, 0.5
                '''                    For I = 1 To RetNofPairs
                '''                        .MoveTo pX(P1(I)), pY(P1(I))
                '''                        .LineTo pX(P2(I)), pY(P2(I))
                '''                        .Stroke
                '''                    Next
                '''                End If

                '--------- DRAW Gravity
                .SetSourceColor vbGreen
                .Arc 30, 30, 25
                .Stroke
                .MoveTo 30, 30
                .LineTo 30 + gX * 25 * InvGravScale, 30 + gY * 25 * InvGravScale
                .Stroke
                .Arc 30 + gX * 25 * InvGravScale, 30 + gY * 25 * InvGravScale, 15 * Abs(gZ) * InvGravScale
                If gZ > 0 Then .Fill Else: .Stroke


                If COMGravity Then
                    .DrawRegularPolygon 30, 80, 20, 8, splNormal, 5
                    .Stroke
                End If
                '---------


                .SelectFont "Courier New", 10, vbGreen, True
                .TextOut 80, 4, "Pts: " & Format$(NP, "###,###,###") & "    " & "h = " & H & "   Pairs: " & OnP & "   "
                .TextOut WW - 185, 4, "Simple SPH by miorsoft"


                '                .TextOut 80, 35, "  camera Pos X   " & CAMERA.Position.X
                '                .TextOut 80, 55, "             Y   " & CAMERA.Position.Y
                '                .TextOut 80, 75, "             Z   " & CAMERA.Position.Z
                '                CAMERA.GetRotation yyy, ppp
                '                .TextOut 80, 95, "            Yaw   " & yyy
                '                .TextOut 80, 115, "            Pitch " & ppp


            End With


            SRF.DrawToDC PIChDC
            If DoSaveFrames Then SRF.WriteContentToPngFile App.Path & "\Frames\" & Format$(Frame, "0000") & ".png": Frame = Frame + 1

        End If


        CNT = CNT + 1&

        If rndGravity Then
            '            If (CNT And 511) = 0& Then
            If (CNT Mod 900) = 0& Then


                gTOX = (Rnd * 2 - 1) * GravScale
                gTOY = (Rnd * 2 - 1) * GravScale
                gTOZ = (Rnd * 2 - 1) * GravScale * 1.1
                If Rnd > 0.2 Then  '0.1
                    If Abs(gTOX) > Abs(gTOY) And Abs(gTOX) > Abs(gTOZ) Then gTOY = 0: gTOZ = 0
                    If Abs(gTOY) > Abs(gTOX) And Abs(gTOY) > Abs(gTOX) Then gTOX = 0: gTOZ = 0
                    If Abs(gTOZ) > Abs(gTOY) And Abs(gTOZ) > Abs(gTOX) Then gTOY = 0: gTOX = 0

                    If Rnd < 0.1 Then gTOX = 0: gTOY = 0: gTOZ = 0
                End If
            End If

            Line1.X1 = picGravity.Width * 0.5
            Line1.Y1 = picGravity.Height * 0.5
            Line1.X2 = Line1.X1 + gTOX * picGravity.Width * 0.5 * InvGravScale
            Line1.Y2 = Line1.Y1 + gTOY * picGravity.Height * 0.5 * InvGravScale

        End If



        gX = gX * 0.98 + gTOX * 0.02
        gY = gY * 0.98 + gTOY * 0.02
        gZ = gZ * 0.98 + gTOZ * 0.02


        ''''CameraMoveWithGravity
        '''CAMERA.Position = SUM3(CAMERA.lookat, MUL3(Normalize3(Vec3(gZ, -Abs(gX), gY)), 680))
        ''tV = SUM3(MUL3(DIFF3(CAMERA.Position, CAMERA.lookat), 0.95), MUL3(Normalize3(Vec3(gZ, -Abs(gX), gY)), 0.05 * 610))
        ''CAMERA.Position = SUM3(CAMERA.lookat, tV)
        '''CAMERA.VectorUP = SUM3(MUL3(CAMERA.VectorUP, 0.5), MUL3(Vec3(-gX, -gY, -gZ), 0.5)): RecomputeBOX = True
        ''CAMERA.VectorUP = Vec3(-gX, -gY, -gZ): RecomputeBOX = True


        If DoFaucet1 Then FaucetSource (1)
        If DoFaucet2 Then FaucetSource (2)
        'If DoFaucet1 Then FaucetSourceB (1)
        'If DoFaucet2 Then FaucetSourceB (2)


        If (CNT And 31&) = 0& Then
            fMain.Caption = "NP: " & NP & "     Pairs: " & RetNofPairs & "    computed FPS: " & FPS & "    SortSwaps: " & SORTSWAPS & "       MaxDensity: " & TestMaxDens
            OnP = Format$(RetNofPairs, "###,###,###")
        End If



        If CamRot Then
            RecomputeBOX = True
            CAMERA.SetPositionAndLookAt Vec3(WW * 0.5 + Cos(CNT * 0.0007) * 520, HH * 0.5, ZZ * 0.5 + Sin(CNT * 0.0007) * 520), Vec3(WW * 0.5, HH * 0.5, ZZ * 0.5)
        End If


        FauxDoEvents

    Loop While DoLOOP

    '    Set SpatialGRID = Nothing
    Set OCTTREE = Nothing

End Sub




Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then

        CAMERA.SetRotation 0, 0

        CAMERA.Position = SUM3(MUL3(Normalize3(DIFF3(CAMERA.Position, CAMERA.lookat)), -WW * 0.7), CAMERA.lookat)

        RecomputeBOX = True
    End If

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D         As Single
    Dim Pitch     As Single
    Dim Yaw       As Single


    Static X0!, Y0!, DX!, DY!

    DX = X - X0: DY = Y - Y0


    Select Case Button
        Case 0
            X0 = X: Y0 = Y
        Case 1

            CAMERA.GetRotation Yaw, Pitch

            'Left hand
            Pitch = Pitch - 0.25 * DY
            Yaw = (Yaw + 0.25 * DX)
            '        Right Hand
            '        Pitch = Pitch - 0.25 * DY
            '        Yaw = (Yaw + 0.25 * dx)

            X0 = X: Y0 = Y
            '        If Pitch > 90 Then Pitch = 90
            '        If Pitch < -90 Then Pitch = -90
            CAMERA.SetRotation Yaw, Pitch

            RecomputeBOX = True

        Case 2                     'zoom
            D = Length3(DIFF3(CAMERA.Position, CAMERA.lookat))
            D = D - DY * 0.25
            '            If D < WW * 0.7 Then D = WW * 0.7

            With CAMERA
                .Position = SUM3(MUL3(Normalize3(DIFF3(.Position, .lookat)), D), .lookat)
            End With

            RecomputeBOX = True

    End Select

End Sub

Private Sub picGravity_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then

        X = (X \ 5) * 5
        Y = (Y \ 5) * 5

        gTOX = 2 * (X - picGravity.Width * 0.5) / picGravity.Width * GravScale
        gTOY = 2 * (Y - picGravity.Height * 0.5) / picGravity.Height * GravScale

        Line1.X1 = picGravity.Width * 0.5
        Line1.Y1 = picGravity.Height * 0.5
        Line1.X2 = X
        Line1.Y2 = Y

    ElseIf Button = 2 Then
        gTOX = 0
        gTOY = 0

        Line1.X1 = picGravity.Width * 0.5
        Line1.Y1 = picGravity.Height * 0.5
        Line1.X2 = Line1.X1
        Line1.Y2 = Line1.Y1

    End If


    gTOZ = 0


End Sub

Private Sub ScrollDRAW_Change()
    RenderEvery = 2 ^ ScrollDRAW.Value - 1
    lSkip = "Skip Render: " & RenderEvery & " Frames"
End Sub

Private Sub ScrollDRAW_Scroll()
    RenderEvery = 2 ^ ScrollDRAW.Value - 1
    lSkip = "Skip Render: " & RenderEvery & " Frames"
End Sub

Private Sub Timer1_Timer()
    ' FPS = CNT - OldCNT: OldCNT = CNT

    OldmTime = mTime
    mTime = Timer

    FPS = Int(100 * (CNT - OldCNT) / (mTime - OldmTime)) * 0.01
    OldCNT = CNT
End Sub


Private Sub FaucetSource(fPhase As Long)
    Dim A         As Single
    Dim C         As Single
    Dim S         As Single
    Dim X         As Single
    Dim Y         As Single
    Dim L         As Single
    Dim sX        As Single
    Dim sY        As Single

    A = (CNT * 0.0125)


    If fPhase = 1 Then
        sX = WW * 1 / 3
        sY = HH * 1 / 3
    Else
        sX = WW * 2 / 3
        sY = HH * 1 / 3
        A = -A                     '+ 3.14159265358979 '* 0.5
    End If

    C = Cos(A)
    S = Sin(A)

    For L = -H * 1 To H * 1 Step H * 0.25
        X = sX + C * L
        Y = sY + S * L
        NP = NP + 1

        ReDim Preserve px(NP)
        ReDim Preserve py(NP)
        ReDim Preserve pz(NP)

        ReDim Preserve vX(NP)
        ReDim Preserve vY(NP)
        ReDim Preserve vZ(NP)

        ReDim pvX(NP)
        ReDim pvY(NP)
        ReDim pvZ(NP)


        ReDim Preserve VXChange(NP)
        ReDim Preserve VYChange(NP)
        ReDim Preserve VZChange(NP)

        ReDim Preserve Density(NP)
        ReDim Preserve INVDensity(NP)
        ReDim Preserve Pressure(NP)
        ReDim Preserve Phase(NP)

        px(NP) = X
        py(NP) = Y
        pz(NP) = ZZ * 0.5

        vX(NP) = -S * H * invDT * 0.25
        vY(NP) = C * H * invDT * 0.25
        vZ(NP) = Rnd * 0.01 * H * invDT

        Phase(NP) = fPhase

    Next

End Sub



Private Sub FaucetSourceB(fPhase As Long)

    Dim X         As Single
    Dim Y         As Single
    Dim sX        As Single
    Dim sY        As Single

    Dim R As Single, stp As Single



    If fPhase = 1 Then
        sX = WW * 1 / 3
        sY = ZZ * 0.5              ' HH * 1 / 3
    Else
        sX = WW * 2 / 3
        sY = ZZ * 0.5              'HH * 1 / 3
    End If

    R = H * 1.2
    stp = H * 0.5

    For Y = sY - R To sY + R Step stp
        For X = sX - R To sX + R Step stp

            If (X - sX) ^ 2 + (Y - sY) ^ 2 <= R * R Then
                NP = NP + 1

                ReDim Preserve px(NP)
                ReDim Preserve py(NP)
                ReDim Preserve pz(NP)

                ReDim Preserve vX(NP)
                ReDim Preserve vY(NP)
                ReDim Preserve vZ(NP)

                ReDim pvX(NP)
                ReDim pvY(NP)
                ReDim pvZ(NP)


                ReDim Preserve VXChange(NP)
                ReDim Preserve VYChange(NP)
                ReDim Preserve VZChange(NP)

                ReDim Preserve Density(NP)
                ReDim Preserve INVDensity(NP)
                ReDim Preserve Pressure(NP)
                ReDim Preserve Phase(NP)

                px(NP) = X + Rnd * H * 0.2
                py(NP) = HH * 0.1 + Rnd * H * 0.2
                pz(NP) = Y + Rnd * H * 0.2

                vX(NP) = 0
                vY(NP) = H * invDT * 0.35    ' 0.25
                vZ(NP) = 0

                Phase(NP) = fPhase
            End If
        Next
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DoLOOP = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    New_c.CleanupRichClientDll
End Sub


Private Sub OLDPointToScreen(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                             rX As Single, rY As Single, rZ As Single)
    rZ = 0.5 + 0.5 * Z * invZZ
    rX = WW * 0.5 + (X - WW * 0.5) * 0.9 * rZ
    rY = HH * 0.5 + (Y - HH * 0.5) * 0.9 * rZ
End Sub

