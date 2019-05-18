Attribute VB_Name = "m3DEasyCam"
' NEEDS mVectors.bas

Option Explicit

Private Const Epsilon As Double = 0.001
Private Const Deg2Rad As Double = 1.74532925199433E-02     'Degrees to Radians
Private Const Rad2Deg As Double = 57.2957795130823      'Radians to Degrees

Public Type tCamera
    cFrom         As tVec3    'Camera Position
    cTo           As tVec3    'Camera LookAT
    ScreenCenter  As tVec3    'Center coords of screen
    camUU         As tVec3    'Cam Matrix
    camVV         As tVec3
    camWW         As tVec3
    VectorUP      As tVec3    'Vector UP

    NearPlaneDist As Double
End Type

Public Camera     As tCamera  'Camera V4

Public Pitch      As Double
Public Yaw        As Double



Public Sub CameraInit(CameraFrom As tVec3, CameraTo As tVec3, ScreenCenter As tVec3, UP As tVec3)
    Camera.cFrom = CameraFrom
    Camera.cTo = CameraTo
    Camera.ScreenCenter = ScreenCenter
    Camera.VectorUP = UP

    Camera.NearPlaneDist = 5

    CameraUpdate
End Sub


Public Sub CameraUpdate()
' Call this every time you change Camera Position or Target !!!
'    // camera matrix
    Dim dx        As Double
    Dim DY        As Double
    Dim dz        As Double



    With Camera
        .camWW = Normalize3(DIFF3(.cTo, .cFrom))
        .camUU = Normalize3(CROSS3(.camWW, .VectorUP))
        .camVV = Normalize3(CROSS3(.camUU, .camWW))
    End With


    '    'cameraUP = Z

    'dx = Camera.cFrom.X - Camera.cTo.X
    'dy = Camera.cFrom.Y - Camera.cTo.Y
    'dz = Camera.cFrom.Z - Camera.cTo.Z
    '
    '    Pitch = Rad2Deg * Atan2(Sqr(dx * dx + dy * dy), dz)  'OK
    '    Yaw = Rad2Deg * -Atan2(-dx, dz)

End Sub

Public Sub CameraSetRotation(ByVal Yaw As Double, ByVal Pitch As Double)



    Dim D         As Double
    ' Thanks to Passel:
    ' http://www.vbforums.com/showthread.php?870755-3D-Swimming-Fish-Algorithm&p=5356667&viewfull=1#post5356667

    '    If Pitch > 90 Then Pitch = 90
    '    If Pitch < -90 Then Pitch = -90

    ' Camera UP = Y

    '    With Camera
    '        D = Length3(DIFF3(.cFrom, .cTo))
    '        .cFrom.X = .cTo.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    '        .cFrom.Y = .cTo.Y + D * (Sin(Pitch * Deg2Rad))
    '        .cFrom.Z = .cTo.Z + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    '    End With

    'cameraUP = Z

    With Camera
        D = Length3(DIFF3(.cFrom, .cTo))
        .cFrom.x = .cTo.x + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
        .cFrom.y = .cTo.y + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
        .cFrom.Z = .cTo.Z + D * (Sin(Pitch * Deg2Rad))
    End With

    CameraUpdate

End Sub
Public Function PointToScreenWDCam(WorldPos As tVec3, ProjectedDistFromCam As Double) As tVec3
    Dim P         As tVec3
    Dim S         As tVec3
    Dim IZ        As Double

    S = WorldPos

    With Camera

        S = DIFF3(S, .cFrom)

        P.x = DOT3(S, .camUU)
        P.y = DOT3(S, .camVV)
        P.Z = DOT3(S, .camWW)

        IZ = 1 / P.Z
        PointToScreenWDCam.x = P.x * IZ * .ScreenCenter.x + .ScreenCenter.x
        PointToScreenWDCam.y = P.y * IZ * .ScreenCenter.x + .ScreenCenter.y
        PointToScreenWDCam.Z = IZ  ' if its negative point is behind camera
        ProjectedDistFromCam = P.Z  ' if its negative point is behind camera

    End With

End Function

Public Sub PointToScreenCoords(ByVal x As Double, ByVal y As Double, ByVal Z As Double, _
                               rX As Double, rY As Double, rZ As Double)

    Dim S         As tVec3
    Dim P         As tVec3
    Dim IZ        As Double

    S = Vec3(x, y, Z)
    With Camera
        S = DIFF3(S, .cFrom)

        P.x = DOT3(S, .camUU)
        P.y = DOT3(S, .camVV)
        P.Z = DOT3(S, .camWW)

        IZ = 1 / P.Z

        rX = P.x * IZ * .ScreenCenter.x + .ScreenCenter.x
        rY = P.y * IZ * .ScreenCenter.x + .ScreenCenter.y
        rZ = IZ  ' if its negative point is behind camera

    End With

End Sub

Public Sub LineToScreen(P1 As tVec3, P2 As tVec3, Ret1 As tVec3, Ret2 As tVec3)

    Dim PlaneCenter As tVec3
    Dim PlaneNormal As tVec3
    Dim IntersectP1 As tVec3
    Dim IntersectP2 As tVec3
    Dim DfromCam1 As Double
    Dim DfromCam2 As Double

    Ret1 = PointToScreenWDCam(P1, DfromCam1)
    Ret2 = PointToScreenWDCam(P2, DfromCam2)

    If DfromCam1 < Camera.NearPlaneDist Then

        If DfromCam2 < Camera.NearPlaneDist Then Exit Sub    'Both points behind camera so EXIT

        'Just P1 Behind, So Find it's intersection To Near plane
        PlaneNormal = Camera.camWW
        PlaneCenter = SUM3(Camera.cFrom, MUL3(PlaneNormal, Camera.NearPlaneDist))
        IntersectP1 = RayPlaneIntersect(DIFF3(P2, P1), P1, PlaneNormal, PlaneCenter)
        Ret1 = PointToScreenWDCam(IntersectP1, DfromCam1)


    ElseIf DfromCam2 < Camera.NearPlaneDist Then

        If DfromCam1 < Camera.NearPlaneDist Then Exit Sub    'Both points behind camera so EXIT


        'Just P2 Behind, So Find it's intersection To Near plane
        PlaneNormal = Camera.camWW
        PlaneCenter = SUM3(Camera.cFrom, MUL3(PlaneNormal, Camera.NearPlaneDist))
        IntersectP2 = RayPlaneIntersect(DIFF3(P1, P2), P2, PlaneNormal, PlaneCenter)
        Ret2 = PointToScreenWDCam(IntersectP2, DfromCam2)

    End If


End Sub



Public Function IsPointVisible(x As Double, y As Double, Z As Double) As Boolean

    If Z < 0 Then Exit Function

    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > Camera.ScreenCenter.x * 2 Then Exit Function
    If y > Camera.ScreenCenter.y * 2 Then Exit Function
    IsPointVisible = True

End Function
