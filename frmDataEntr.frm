VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataEntr 
   Caption         =   "Employee Registration Form"
   ClientHeight    =   6276
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5688
   OleObjectBlob   =   "frmDataEntr.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDataEntr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReset_Click()

   Dim i As VbMsgBoxResult
   i = MsgBox("¿Quieres limpiar el formulario?", vbYesNo + vbQuestion, "Limpiar formulario")
   If i = vbNo Then Exit Sub
   Call Reset_Form

End Sub

Private Sub cmdSubmit_Click()
Dim i As VbMsgBoxResult

i = MsgBox("¿Quieres enviar la informacion?", vbYesNo + vbQuestion, "Enviar informacion")

If i = vbNo Then Exit Sub

If ValidEntry Then

    Call Submit_Data

End If
End Sub


Private Sub UserForm_Initialize()
    Call Reset_Form
End Sub

