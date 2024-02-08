Attribute VB_Name = "mFauxDoEvents"
Option Explicit


Private Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long

Private Const QS_KEY As Long = &H1
Private Const QS_MOUSEBUTTON As Long = &H4
Attribute QS_MOUSEBUTTON.VB_VarUserMemId = 1073938434
Private Const QS_POSTMESSAGE As Long = &H8
Attribute QS_POSTMESSAGE.VB_VarUserMemId = 1073938435
Private Const QS_SENDMESSAGE As Long = &H40

Private Const EveryThingsOR As Long = 77 + 2    'reexre
Attribute EveryThingsOR.VB_VarUserMemId = 1610809344

Public Sub FauxDoEvents()
Attribute FauxDoEvents.VB_UserMemId = 1073741825
    ' pulled from this posting
    ' http://www.vbforums.com/showthread.php?315416-Ok-noobies-fauxdoevents-is-slow!!!-Here-s-are-faster-methods

    ' only calls DoEvents when absolutely necessary.
    ' potential side-effect: if form is marked by Windows as "Not Responding",
    '   it should clear relatively quickly but in doing so, form visibly repaints


    '   If GetQueueStatus(QS_KEY Or QS_MOUSEBUTTON Or QS_POSTMESSAGE Or QS_SENDMESSAGE) <> 0 Then DoEvents

    If GetQueueStatus(EveryThingsOR) Then DoEvents


End Sub


