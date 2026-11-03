Attribute VB_Name = "modHover"
Option Explicit

' ============================================================
' modHover - remembers which sidebar button is highlighted
' ============================================================

Private mCurrent As MSForms.CommandButton
Private mNormalColor As Long
Private mHoverColor As Long

' Highlights a button (and un-highlights the previous one).
Public Sub SetHover(ByVal btn As MSForms.CommandButton, ByVal normalColor As Long, _
                    ByVal hoverColor As Long)
    If Not mCurrent Is Nothing Then
        If mCurrent Is btn Then Exit Sub            ' already highlighted: nothing to do
        ClearHover
    End If

    ' Only highlight an enabled button that is in its normal colour
    If btn.Enabled Then
        If btn.BackColor = normalColor Then
            btn.BackColor = hoverColor
            Set mCurrent = btn
            mNormalColor = normalColor
            mHoverColor = hoverColor
        End If
    End If
End Sub

' Removes the highlight.
Public Sub ClearHover()
    If mCurrent Is Nothing Then Exit Sub

    On Error Resume Next
    ' Only undo the highlight if nothing else changed the colour meanwhile
    If mCurrent.BackColor = mHoverColor Then mCurrent.BackColor = mNormalColor
    On Error GoTo 0

    Set mCurrent = Nothing
End Sub

