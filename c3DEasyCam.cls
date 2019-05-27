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
Private Const Deg2Rad As Double = 1.74532925199433E-02     'Degrees to Radians
Private Const Rad2Deg As Double = 57.2957795130823      'Radians to Degrees

Private mPosition As tVec3    'Camera Position
Private mLookAt   As tVec3    'Camera LookAT
Private mScreenCenter As tVec3    'Center coords of screen
Private camUU     As tVec3    'Cam Matrix
Private camVV     As tVec3
Private camWW     As tVec3
Private mVectorUP As tVec3    'Vector UP
Private NearPlaneDist As Double

Private mPitch    As Double
Private mYaw      As Double
Private mZOOM     As Double


Friend Property Let Position(V As tVec3)
    mPosition = V
    Update
End Property

Friend Property Get Position() As tVec3
    Position = mPosition
End Property

Friend Property Let lookat(V As tVec3)
    mLookAt = V
    Update
End Property

Friend Property Get lookat() As tVec3
    lookat = mLookAt
End Property

Friend Sub SetPositionAndLookAt(Pos As tVec3, Look As tVec3)
    mPosition = Pos
    mLookAt = Look
    Update
End Sub

Friend Property Get Direction() As tVec3
    Direction = camWW
End Property


Friend Property Let VectorUP(UP As tVec3)
    mVectorUP = Normalize3(UP)
    Update
End Property

Friend Property Get VectorUP() As tVec3
    VectorUP = mVectorUP
End Property


Friend Sub GetRotation(ByRef Yaw As Double, ByRef Pitch As Double)
    Update True
    Pitch = mPitch
    Yaw = mYaw
End Sub

Friend Property Let Zoom(V As Double)
    mZOOM = V
End Property

Friend Property Get Zoom() As Double
    Zoom = mZOOM
End Property



Friend Sub Init(CameraFrom As tVec3, CameraTo As tVec3, ScreenCenter As tVec3, UP As tVec3)
    mPosition = CameraFrom
    mLookAt = CameraTo
    mScreenCenter = ScreenCenter
    mVectorUP = UP

    NearPlaneDist = 5

    mZOOM = 1

    Update
End Sub


Private Sub Update(Optional UpdatePitchYawValues As Boolean = False)
' Call this every time you change Camera Position or Target !!!
'    // camera matrix
    Dim D         As tVec3

' Left Hand
'    camWW = Normalize3(DIFF3(mLookAt, mPosition))
'    camUU = Normalize3(CROSS3(mVectorUP, camWW))
'    camVV = Normalize3(CROSS3(camUU, camWW))


' Right Hand
    camWW = Normalize3(DIFF3(mLookAt, mPosition))
    camUU = Normalize3(CROSS3(camWW, mVectorUP))
    camVV = Normalize3(CROSS3(camUU, camWW))


    If UpdatePitchYawValues Then    '  (GetRotation)
        ' ############    CAMERA mVectorUP  MUST BE 0,-1,0
        'https://math.stackexchange.com/questions/470112/calculate-camera-pitch-yaw-to-face-point
        'D = DIFF3(Position, lookat)
        
        'Left Hand
'        D = camWW    'opposite sign than above
'        With D
'            mPitch = (-Atan2(Sqr(.X * .X + .Z * .Z), .Y)) * Rad2Deg
'            mYaw = (-Atan2(-.X, -.Z) + PIh) * Rad2Deg
'        End With
        
        
        'Right Hand
         D = camWW    'opposite sign than above
        With D
            mPitch = (Atan2(Sqr(.X * .X + .Z * .Z), -.Y)) * Rad2Deg
            mYaw = (-Atan2(-.X, -.Z) + PIh) * Rad2Deg
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
'    D = Length3(DIFF3(mPosition, mLookAt))
'    mPosition.X = mLookAt.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
'    mPosition.Y = mLookAt.Y + D * (Sin(Pitch * Deg2Rad))
'    mPosition.Z = mLookAt.Z + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))


'    '    'cameraUP = Z
'            D = Length3(DIFF3(.mPosition, .mLookAt))
'            mPosition.X = mLookAt.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
'            mPosition.Y = mLookAt.Y + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
'            mPosition.Z = mLookAt.Z + D * (Sin(Pitch * Deg2Rad))


    D = Length3(DIFF3(mPosition, mLookAt))
    mPosition.X = mLookAt.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    mPosition.Y = mLookAt.Y + D * (Sin(Pitch * Deg2Rad))
    mPosition.Z = mLookAt.Z + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))


    Update

End Sub

Friend Function PointToScreenWDCam(WorldPos As tVec3, ProjectedDistFromCam As Double) As tVec3
    Dim p         As tVec3
    Dim S         As tVec3
    Dim IZ        As Double

    S = DIFF3(WorldPos, mPosition)
    p.X = DOT3(S, camUU)
    p.Y = DOT3(S, camVV)
    p.Z = DOT3(S, camWW)
    IZ = 1 / p.Z * mZOOM
    PointToScreenWDCam.X = p.X * IZ * mScreenCenter.X + mScreenCenter.X
        'left hand
'PointToScreenWDCam.Y = p.Y * IZ * mScreenCenter.X + mScreenCenter.Y
        'Right hand
    PointToScreenWDCam.Y = -p.Y * IZ * mScreenCenter.X + mScreenCenter.Y
    
    PointToScreenWDCam.Z = IZ    ' if its negative point is behind camera
    ProjectedDistFromCam = p.Z  ' if its negative point is behind camera

End Function

Friend Sub PointToScreenCoords(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, _
                               rX As Double, rY As Double, rZ As Double)
    Dim S         As tVec3
    Dim p         As tVec3
    Dim IZ        As Double

    S = Vec3(X, Y, Z)
    S = DIFF3(S, mPosition)
    p.X = DOT3(S, camUU)
    p.Y = DOT3(S, camVV)
    p.Z = DOT3(S, camWW)
    IZ = 1 / p.Z * mZOOM
    rX = p.X * IZ * mScreenCenter.X + mScreenCenter.X
    'left hand
    'rY = p.Y * IZ * mScreenCenter.X + mScreenCenter.Y
    'Right hand
    rY = -p.Y * IZ * mScreenCenter.X + mScreenCenter.Y
    
    rZ = IZ    ' if its negative point is behind camera

End Sub

Friend Sub LineToScreen(P1 As tVec3, P2 As tVec3, Ret1 As tVec3, Ret2 As tVec3)

    Dim PlaneCenter As tVec3
    Dim PlaneNormal As tVec3
    Dim IntersectP1 As tVec3
    Dim IntersectP2 As tVec3
    Dim DfromCam1 As Double
    Dim DfromCam2 As Double

    Ret1 = PointToScreenWDCam(P1, DfromCam1)
    Ret2 = PointToScreenWDCam(P2, DfromCam2)

    If DfromCam1 < NearPlaneDist Then

        If DfromCam2 < NearPlaneDist Then Exit Sub    'Both points behind camera so EXIT

        'Just P1 Behind, So Find it's intersection To Near plane
        PlaneNormal = camWW
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, NearPlaneDist))
        IntersectP1 = RayPlaneIntersect(DIFF3(P2, P1), P1, PlaneNormal, PlaneCenter)
        Ret1 = PointToScreenWDCam(IntersectP1, DfromCam1)

    ElseIf DfromCam2 < NearPlaneDist Then

        If DfromCam1 < NearPlaneDist Then Exit Sub    'Both points behind camera so EXIT

        'Just P2 Behind, So Find it's intersection To Near plane
        PlaneNormal = camWW
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, NearPlaneDist))
        IntersectP2 = RayPlaneIntersect(DIFF3(P1, P2), P2, PlaneNormal, PlaneCenter)
        Ret2 = PointToScreenWDCam(IntersectP2, DfromCam2)

    End If

End Sub


Friend Function IsPointVisibleCoords(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Boolean

    If Z < 0 Then Exit Function    ' behind

    If X < 0 Then Exit Function    ' check if outside screen
    If Y < 0 Then Exit Function
    If X > mScreenCenter.X * 2 Then Exit Function
    If Y > mScreenCenter.Y * 2 Then Exit Function

    IsPointVisibleCoords = True

End Function


Friend Function IsPointVisible(V As tVec3) As Boolean

    With V
        If .Z < 0 Then Exit Function    ' behind

        If .X < 0 Then Exit Function    ' check if outside screen
        If .Y < 0 Then Exit Function
        If .X > mScreenCenter.X * 2 Then Exit Function
        If .Y > mScreenCenter.Y * 2 Then Exit Function
    End With

    IsPointVisible = True

End Function

Friend Function IsPointVisibleGap(V As tVec3, Gap As Double) As Boolean

    With V
        If .Z < 0 Then Exit Function    ' behind

        If .X < -Gap Then Exit Function    ' check if outside screen
        If .Y < -Gap Then Exit Function
        If .X - Gap > mScreenCenter.X * 2 Then Exit Function
        If .Y - Gap > mScreenCenter.Y * 2 Then Exit Function
    End With

    IsPointVisibleGap = True

End Function

Friend Function Follow(Target As tVec3, LookAtSpeed01 As Double, PositionSpeed01 As Double, MAXDist2 As Double, Optional MINDist2 As Double = 0, Optional PosMinY As Double = 0)
    Dim D         As Double


    mLookAt = SUM3(MUL3(mLookAt, 1 - LookAtSpeed01), MUL3(Target, LookAtSpeed01))

    D = Length32(DIFF3(mLookAt, mPosition))
    If D > MAXDist2 Then
        mPosition = SUM3(MUL3(mPosition, 1 - PositionSpeed01), MUL3(SUM3(Target, MUL3(camWW, -Sqr(MAXDist2))), PositionSpeed01))
    ElseIf D < MINDist2 Then
        mPosition = SUM3(MUL3(mPosition, 1 - PositionSpeed01), MUL3(SUM3(Target, MUL3(camWW, -Sqr(MINDist2))), PositionSpeed01))
    End If

'Left Hand:
'    If PosMinY Then If mPosition.Y > PosMinY Then mPosition.Y = PosMinY
'Right Hand:
    If PosMinY Then If mPosition.Y < PosMinY Then mPosition.Y = PosMinY


    Update

End Function

