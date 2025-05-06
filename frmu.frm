VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmu 
   Caption         =   "Discover Shared Drive"
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3750
   OleObjectBlob   =   "frmu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
mmmppp

End Sub

Private Sub UserForm_INITIALIZE()

Me.cmba.List = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")


End Sub


Private Sub mmmppp()
Dim dlett As String
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("Wscript.Network")

dlett = Trim(Me.cmba.Value)
dlett = dlett & ":"

If (objFSO.DriveExists(dlett) = True) Then
    objNetwork.RemoveNetworkDrive dlett, True, True
End If

objNetwork.MapNetworkDrive dlett, "\\PL1GVFS0011\AAA_Shared"

MsgBox "\\PL1GVFS0011\AAA_Shared is now mapped to " & dlett

Unload Me


End Sub


