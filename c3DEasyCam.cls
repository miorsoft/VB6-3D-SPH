VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c3DEasyCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' NEEDS mVectors.bas

' 3D CAMERA by
' reexre - miorsoft - Roberto Mior

' WORKS ONLY with mVectorUP = 0, -1 , 0

' LEFT HAND RULES
' https://upload.wikimedia.org/wikipedia/commons/b/b2/3D_Cartesian_Coodinate_Handedness.jpg



Option Explicit

Private Const Epsilon As Single = 0.001
Attribute Epsilon.VB_VarUserMemId = 1610809345
Private Const Deg2Rad As Single = 1.74532925199433E-02    'Degrees to Radians
Attribute Deg2Rad.VB_VarUserMemId = 1073938435
Private Const Rad2Deg As Single = 57.2957795130823    'Radians to Degrees
Attribute Rad2Deg.VB_VarUserMemId = 1073938436

Private mPosition As tVec3         'Camera Position
Attribute mPosition.VB_VarUserMemId = 1073938438
Private mLookAt   As tVec3         'Camera LookAT
Attribute mLookAt.VB_VarUserMemId = 1073938433
Private mScreenCenter As tVec3     'Center coords of screen
Attribute mScreenCenter.VB_VarUserMemId = 1073741830
Private camRIGHT  As tVec3         'Cam Matrix
Attribute camRIGHT.VB_VarUserMemId = 1610809346
Private camUP     As tVec3
Attribute camUP.VB_VarUserMemId = 1073938440
Private camFRONT  As tVec3
Attribute camFRONT.VB_VarUserMemId = 1073938441
Private mVectorUP As tVec3         'Vector UP
Attribute mVectorUP.VB_VarUserMemId = 1073938442
Private mNearPlaneDist As Single
Attribute mNearPlaneDist.VB_VarUserMemId = 1073938443
Private mFarPlaneDist As Single
Attribute mFarPlaneDist.VB_VarUserMemId = 1073938444


Private mPitch    As Single
Attribute mPitch.VB_VarUserMemId = 1073741837
Private mYaw      As Single
Attribute mYaw.VB_VarUserMemId = 1073741838
Private mZOOM     As Single
Attribute mZOOM.VB_VarUserMemId = 1610809348


Friend Property Let Position(V As tVec3)
    mPosition = V
    UPDATE
End Property

Friend Property Get Position() As tVec3
    Position = mPosition
End Property

Friend Property Let lookat(V As tVec3)
    mLookAt = V
    UPDATE
End Property

Friend Property Get lookat() As tVec3
    lookat = mLookAt
End Property

Friend Sub SetPositionAndLookAt(Pos As tVec3, Look As tVec3)
    mPosition = Pos
    mLookAt = Look
    UPDATE
End Sub

Friend Property Get Direction() As tVec3
    Direction = camFRONT
End Property


Friend Property Let VectorUP(UP As tVec3)
    mVectorUP = Normalize3(UP)
    UPDATE
End Property

Friend Property Get VectorUP() As tVec3
    VectorUP = mVectorUP
End Property


Friend Sub GetRotation(ByRef Yaw As Single, ByRef Pitch As Single)
    UPDATE True
    Pitch = mPitch
    Yaw = mYaw
End Sub

Friend Property Let Zoom(V As Single)
    mZOOM = V
End Property

Friend Property Get Zoom() As Single
    Zoom = mZOOM
End Property

Friend Property Let FarPlane(V As Single)
    mFarPlaneDist = V
End Property

Friend Property Get FarPlane() As Single
    FarPlane = mFarPlaneDist
End Property

Friend Property Let NearPlane(V As Single)
    mNearPlaneDist = V
End Property

Friend Property Get NearPlane() As Single
    NearPlane = mNearPlaneDist
End Property

Friend Sub Init(CameraFrom As tVec3, CameraTo As tVec3, ScreenCenter As tVec3, UP As tVec3)
    mPosition = CameraFrom
    mLookAt = CameraTo
    mScreenCenter = ScreenCenter
    mVectorUP = Normalize3(UP)

    mNearPlaneDist = 5
    mFarPlaneDist = 1E+32

    mZOOM = 1

    UPDATE
End Sub


Private Sub UPDATE(Optional UpdatePitchYawValues As Boolean = False)
    ' Call this every time you change Camera Position or Target !!!
    '    // camera matrix
    Dim D         As tVec3

    ' Left Hand
    '    camFRONT = Normalize3(DIFF3(mLookAt, mPosition))
    '    camRIGHT = Normalize3(CROSS3(mVectorUP, camFRONT))
    '    camUP = Normalize3(CROSS3(camRIGHT, camFRONT))


    ' Right Hand
    camFRONT = Normalize3(DIFF3(mLookAt, mPosition))
    camRIGHT = Normalize3(CROSS3(camFRONT, mVectorUP))
    camUP = Normalize3(CROSS3(camFRONT, camRIGHT))


    'Debug.Print ToString3(camFRONT)
    'Debug.Print ToString3(camRIGHT)
    'Debug.Print ToString3(camUP)




    If UpdatePitchYawValues Then   '  (GetRotation)
        ' ############    CAMERA mVectorUP  MUST BE 0,-1,0
        'https://math.stackexchange.com/questions/470112/calculate-camera-pitch-yaw-to-face-point
        'D = DIFF3(Position, lookat)

        D = camFRONT               'opposite sign than above
        With D
            mPitch = (-Atan2(Sqr(.X * .X + .Z * .Z), .Y)) * Rad2Deg
            mYaw = (-Atan2(-.X, -.Z) + PIh) * Rad2Deg
        End With

    End If

End Sub



'******************************************************************
'   TODO:
'   http://paulbourke.net/geometry/rotate/
'   http://paulbourke.net/geometry/rotate/source.c
'******************************************************************


Public Sub SetRotation(ByVal Yaw As Single, ByVal Pitch As Single)
Attribute SetRotation.VB_UserMemId = 1610809355

    Dim D         As Single
    ' Thanks to Passel:
    ' http://www.vbforums.com/showthread.php?870755-3D-Swimming-Fish-Algorithm&p=5356667&viewfull=1#post5356667

    ' ############    CAMERA mVectorUP MUST BE 0,-1,0
    D = Length3(DIFF3(mPosition, mLookAt))
    mPosition.X = mLookAt.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    mPosition.Y = mLookAt.Y + D * (Sin(Pitch * Deg2Rad))
    mPosition.Z = mLookAt.Z + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))

    '    '    'cameraUP = Z
    '            D = Length3(DIFF3(.mPosition, .mLookAt))
    '            mPosition.X = mLookAt.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    '            mPosition.Y = mLookAt.Y + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    '            mPosition.Z = mLookAt.Z + D * (Sin(Pitch * Deg2Rad))

    UPDATE

End Sub

Friend Function PointToScreenWDCam(WorldPos As tVec3, ProjectedDistFromCam As Single) As tVec3
    Dim P         As tVec3
    Dim S         As tVec3
    Dim IZ        As Single

    S = DIFF3(WorldPos, mPosition)

    P.X = DOT3(S, camRIGHT)
    P.Y = DOT3(S, camUP)
    P.Z = DOT3(S, camFRONT)


    IZ = 1 / P.Z * mZOOM
    PointToScreenWDCam.X = P.X * IZ * mScreenCenter.X + mScreenCenter.X
    PointToScreenWDCam.Y = P.Y * IZ * mScreenCenter.X + mScreenCenter.Y

    PointToScreenWDCam.Z = IZ      ' if its negative point is behind camera
    ProjectedDistFromCam = P.Z     ' if its negative point is behind camera

End Function

Friend Sub PointToScreenCoords(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                               rX As Single, rY As Single, rZ As Single, InvZ As Single)
    Dim S         As tVec3
    Dim P         As tVec3
    '    Dim InvZ        As single

    S = Vec3(X, Y, Z)
    S = DIFF3(S, mPosition)

    With P
        .X = DOT3(S, camRIGHT)
        .Y = DOT3(S, camUP)
        .Z = DOT3(S, camFRONT)

        InvZ = 1 / .Z
        rX = .X * InvZ * mZOOM * mScreenCenter.X + mScreenCenter.X
        rY = .Y * InvZ * mZOOM * mScreenCenter.X + mScreenCenter.Y
        rZ = .Z                    ' if its negative point is behind camera
    End With

End Sub

Friend Sub LineToScreen(P1 As tVec3, P2 As tVec3, Ret1 As tVec3, Ret2 As tVec3, Visible As Boolean)

    Dim PlaneCenter As tVec3
    Dim PlaneNormal As tVec3
    Dim IntersectP1 As tVec3
    Dim IntersectP2 As tVec3
    Dim DfromCam1 As Single
    Dim DfromCam2 As Single

    Ret1 = PointToScreenWDCam(P1, DfromCam1)
    Ret2 = PointToScreenWDCam(P2, DfromCam2)

    Visible = False


    If DfromCam1 < mNearPlaneDist Then

        If DfromCam2 < mNearPlaneDist Then Exit Sub    'Both points behind camera so EXIT

        'Just P1 Behind, So Find it's intersection To Near plane
        PlaneNormal = camFRONT
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, mNearPlaneDist))
        IntersectP1 = RayPlaneIntersect(DIFF3(P2, P1), P1, PlaneNormal, PlaneCenter)
        Ret1 = PointToScreenWDCam(IntersectP1, DfromCam1)

    ElseIf DfromCam2 < mNearPlaneDist Then

        If DfromCam1 < mNearPlaneDist Then Exit Sub    'Both points behind camera so EXIT

        'Just P2 Behind, So Find it's intersection To Near plane
        PlaneNormal = camFRONT
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, mNearPlaneDist))
        IntersectP2 = RayPlaneIntersect(DIFF3(P1, P2), P2, PlaneNormal, PlaneCenter)
        Ret2 = PointToScreenWDCam(IntersectP2, DfromCam2)

    End If


    'Same for FAR PLANE
    If DfromCam1 > mFarPlaneDist Then

        If DfromCam2 > mFarPlaneDist Then Exit Sub    'Both points Too far from camera so EXIT

        'Just P1 Too far from, So Find it's intersection To Far plane
        PlaneNormal = camFRONT
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, mFarPlaneDist))
        IntersectP1 = RayPlaneIntersect(DIFF3(P2, P1), P1, PlaneNormal, PlaneCenter)
        Ret1 = PointToScreenWDCam(IntersectP1, DfromCam1)

    ElseIf DfromCam2 > mFarPlaneDist Then

        If DfromCam1 > mFarPlaneDist Then Exit Sub    'Both points Too far from camera so EXIT

        'Just P2 Too far from, So Find it's intersection To Far plane
        PlaneNormal = camFRONT
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, mFarPlaneDist))
        IntersectP2 = RayPlaneIntersect(DIFF3(P1, P2), P2, PlaneNormal, PlaneCenter)
        Ret2 = PointToScreenWDCam(IntersectP2, DfromCam2)

    End If

    Visible = True

End Sub


Friend Function IsPointVisibleCoords(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Boolean

    If Z < 0 Then Exit Function    ' behind

    If Z > mFarPlaneDist Then Exit Function    ' too far

    If X < 0 Then Exit Function    ' check if outside screen
    If Y < 0 Then Exit Function
    If X > mScreenCenter.X * 2 Then Exit Function
    If Y > mScreenCenter.Y * 2 Then Exit Function

    IsPointVisibleCoords = True

End Function


Friend Function IsPointVisible(V As tVec3) As Boolean

    With V
        If .Z < 0 Then Exit Function    ' behind
        If .Z > mFarPlaneDist Then Exit Function    ' too far

        If .X < 0 Then Exit Function    ' check if outside screen
        If .Y < 0 Then Exit Function
        If .X > mScreenCenter.X * 2 Then Exit Function
        If .Y > mScreenCenter.Y * 2 Then Exit Function
    End With

    IsPointVisible = True

End Function

Friend Function IsPointVisibleGap(V As tVec3, Gap As Single) As Boolean

    With V
        If .Z < 0 Then Exit Function    ' behind
        If .Z - Gap > mFarPlaneDist Then Exit Function    ' too far

        If .X < -Gap Then Exit Function    ' check if outside screen
        If .Y < -Gap Then Exit Function
        If .X - Gap > mScreenCenter.X * 2 Then Exit Function
        If .Y - Gap > mScreenCenter.Y * 2 Then Exit Function
    End With

    IsPointVisibleGap = True

End Function

Friend Function IsPointVisibleGap2(X As Single, Y As Single, Z As Single, Gap As Single) As Boolean

    If Z < 0 Then Exit Function    ' behind
    If Z - Gap > mFarPlaneDist Then Exit Function    ' too far

    If X < -Gap Then Exit Function    ' check if outside screen
    If Y < -Gap Then Exit Function
    If X - Gap > mScreenCenter.X * 2 Then Exit Function
    If Y - Gap > mScreenCenter.Y * 2 Then Exit Function

    IsPointVisibleGap2 = True

End Function


Friend Function Follow(Target As tVec3, LookAtSpeed01 As Single, PositionSpeed01 As Single, MAXDist2 As Single, Optional MINDist2 As Single = 0, Optional PosMinY As Single = 0)
    Dim D         As Single


    mLookAt = SUM3(MUL3(mLookAt, 1 - LookAtSpeed01), MUL3(Target, LookAtSpeed01))

    D = Length32(DIFF3(mLookAt, mPosition))
    If D > MAXDist2 Then
        mPosition = SUM3(MUL3(mPosition, 1 - PositionSpeed01), MUL3(SUM3(Target, MUL3(camFRONT, -Sqr(MAXDist2))), PositionSpeed01))
    ElseIf D < MINDist2 Then
        mPosition = SUM3(MUL3(mPosition, 1 - PositionSpeed01), MUL3(SUM3(Target, MUL3(camFRONT, -Sqr(MINDist2))), PositionSpeed01))
    End If

    If PosMinY Then If mPosition.Y > PosMinY Then mPosition.Y = PosMinY


    UPDATE

End Function

