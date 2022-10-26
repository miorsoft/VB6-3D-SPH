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

Private Const Epsilon As Double = 0.001
Private Const Deg2Rad As Double = 1.74532925199433E-02    'Degrees to Radians
Private Const Rad2Deg As Double = 57.2957795130823    'Radians to Degrees

Private mPosition As tVec3   'Camera Position
Private mLookAt   As tVec3   'Camera LookAT
Private mScreenCenter As tVec3    'Center coords of screen
Private camRIGHT  As tVec3   'Cam Matrix
Private camUP     As tVec3
Private camFRONT  As tVec3
Private mVectorUP As tVec3   'Vector UP
Private mNearPlaneDist As Double
Private mFarPlaneDist As Double


Private mPitch    As Double
Private mYaw      As Double
Private mZOOM     As Double


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


Friend Sub GetRotation(ByRef Yaw As Double, ByRef Pitch As Double)
    UPDATE True
    Pitch = mPitch
    Yaw = mYaw
End Sub

Friend Property Let Zoom(V As Double)
    mZOOM = V
End Property

Friend Property Get Zoom() As Double
    Zoom = mZOOM
End Property

Friend Property Let FarPlane(V As Double)
    mFarPlaneDist = V
End Property

Friend Property Get FarPlane() As Double
    FarPlane = mFarPlaneDist
End Property

Friend Property Let NearPlane(V As Double)
    mNearPlaneDist = V
End Property

Friend Property Get NearPlane() As Double
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




    If UpdatePitchYawValues Then    '  (GetRotation)
        ' ############    CAMERA mVectorUP  MUST BE 0,-1,0
        'https://math.stackexchange.com/questions/470112/calculate-camera-pitch-yaw-to-face-point
        'D = DIFF3(Position, lookat)

        D = camFRONT         'opposite sign than above
        With D
            mPitch = (-Atan2(Sqr(.x * .x + .z * .z), .y)) * Rad2Deg
            mYaw = (-Atan2(-.x, -.z) + PIh) * Rad2Deg
        End With

    End If

End Sub



'******************************************************************
'   TODO:
'   http://paulbourke.net/geometry/rotate/
'   http://paulbourke.net/geometry/rotate/source.c
'******************************************************************


Public Sub SetRotation(ByVal Yaw As Double, ByVal Pitch As Double)

    Dim D         As Double
    ' Thanks to Passel:
    ' http://www.vbforums.com/showthread.php?870755-3D-Swimming-Fish-Algorithm&p=5356667&viewfull=1#post5356667

    ' ############    CAMERA mVectorUP MUST BE 0,-1,0
    D = Length3(DIFF3(mPosition, mLookAt))
    mPosition.x = mLookAt.x + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    mPosition.y = mLookAt.y + D * (Sin(Pitch * Deg2Rad))
    mPosition.z = mLookAt.z + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))

    '    '    'cameraUP = Z
    '            D = Length3(DIFF3(.mPosition, .mLookAt))
    '            mPosition.X = mLookAt.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    '            mPosition.Y = mLookAt.Y + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    '            mPosition.Z = mLookAt.Z + D * (Sin(Pitch * Deg2Rad))

    UPDATE

End Sub

Friend Function PointToScreenWDCam(WorldPos As tVec3, ProjectedDistFromCam As Double) As tVec3
    Dim P         As tVec3
    Dim S         As tVec3
    Dim IZ        As Double

    S = DIFF3(WorldPos, mPosition)

    P.x = DOT3(S, camRIGHT)
    P.y = DOT3(S, camUP)
    P.z = DOT3(S, camFRONT)


    IZ = 1 / P.z * mZOOM
    PointToScreenWDCam.x = P.x * IZ * mScreenCenter.x + mScreenCenter.x
    PointToScreenWDCam.y = P.y * IZ * mScreenCenter.x + mScreenCenter.y

    PointToScreenWDCam.z = IZ    ' if its negative point is behind camera
    ProjectedDistFromCam = P.z    ' if its negative point is behind camera

End Function

Friend Sub PointToScreenCoords(ByVal x As Double, ByVal y As Double, ByVal z As Double, _
                               rX As Double, rY As Double, rZ As Double, InvZ As Double)
    Dim S         As tVec3
    Dim P         As tVec3
    '    Dim InvZ        As Double

    S = Vec3(x, y, z)
    S = DIFF3(S, mPosition)

    With P
        .x = DOT3(S, camRIGHT)
        .y = DOT3(S, camUP)
        .z = DOT3(S, camFRONT)

        InvZ = 1# / .z
        rX = .x * InvZ * mZOOM * mScreenCenter.x + mScreenCenter.x
        rY = .y * InvZ * mZOOM * mScreenCenter.x + mScreenCenter.y
        rZ = .z              ' if its negative point is behind camera
    End With

End Sub

Friend Sub LineToScreen(P1 As tVec3, P2 As tVec3, Ret1 As tVec3, Ret2 As tVec3, Visible As Boolean)

    Dim PlaneCenter As tVec3
    Dim PlaneNormal As tVec3
    Dim IntersectP1 As tVec3
    Dim IntersectP2 As tVec3
    Dim DfromCam1 As Double
    Dim DfromCam2 As Double

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


Friend Function IsPointVisibleCoords(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Boolean

    If z < 0 Then Exit Function    ' behind

    If z > mFarPlaneDist Then Exit Function    ' too far

    If x < 0 Then Exit Function    ' check if outside screen
    If y < 0 Then Exit Function
    If x > mScreenCenter.x * 2 Then Exit Function
    If y > mScreenCenter.y * 2 Then Exit Function

    IsPointVisibleCoords = True

End Function


Friend Function IsPointVisible(V As tVec3) As Boolean

    With V
        If .z < 0 Then Exit Function    ' behind
        If .z > mFarPlaneDist Then Exit Function    ' too far

        If .x < 0 Then Exit Function    ' check if outside screen
        If .y < 0 Then Exit Function
        If .x > mScreenCenter.x * 2 Then Exit Function
        If .y > mScreenCenter.y * 2 Then Exit Function
    End With

    IsPointVisible = True

End Function

Friend Function IsPointVisibleGap(V As tVec3, Gap As Double) As Boolean

    With V
        If .z < 0# Then Exit Function    ' behind
        If .z - Gap > mFarPlaneDist Then Exit Function    ' too far

        If .x < -Gap Then Exit Function    ' check if outside screen
        If .y < -Gap Then Exit Function
        If .x - Gap > mScreenCenter.x * 2 Then Exit Function
        If .y - Gap > mScreenCenter.y * 2 Then Exit Function
    End With

    IsPointVisibleGap = True

End Function

Friend Function IsPointVisibleGap2(x As Double, y As Double, z As Double, Gap As Double) As Boolean

    If z < 0# Then Exit Function    ' behind
    If z - Gap > mFarPlaneDist Then Exit Function    ' too far

    If x < -Gap Then Exit Function    ' check if outside screen
    If y < -Gap Then Exit Function
    If x - Gap > mScreenCenter.x * 2 Then Exit Function
    If y - Gap > mScreenCenter.y * 2 Then Exit Function

    IsPointVisibleGap2 = True

End Function


Friend Function Follow(Target As tVec3, LookAtSpeed01 As Double, PositionSpeed01 As Double, MAXDist2 As Double, Optional MINDist2 As Double = 0, Optional PosMinY As Double = 0)
    Dim D         As Double


    mLookAt = SUM3(MUL3(mLookAt, 1 - LookAtSpeed01), MUL3(Target, LookAtSpeed01))

    D = Length32(DIFF3(mLookAt, mPosition))
    If D > MAXDist2 Then
        mPosition = SUM3(MUL3(mPosition, 1 - PositionSpeed01), MUL3(SUM3(Target, MUL3(camFRONT, -Sqr(MAXDist2))), PositionSpeed01))
    ElseIf D < MINDist2 Then
        mPosition = SUM3(MUL3(mPosition, 1 - PositionSpeed01), MUL3(SUM3(Target, MUL3(camFRONT, -Sqr(MINDist2))), PositionSpeed01))
    End If

    If PosMinY Then If mPosition.y > PosMinY Then mPosition.y = PosMinY


    UPDATE

End Function

