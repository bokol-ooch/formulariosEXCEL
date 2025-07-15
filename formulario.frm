VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formulario 
   Caption         =   "Registro de Faltantes de Caja"
   ClientHeight    =   8076
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17568
   OleObjectBlob   =   "formulario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub calendario_click()
Dim dateVariable As Date
dateVariable = CalendarForm.GetDate
formulario.txtDia.value = Format(dateVariable, "dd/mmm/yyyy")
End Sub

Private Sub cmdDelete_Click()
If Selected_List = 0 Then

     MsgBox "No hay renglón seleccionado", vbOKOnly + vbInformation, "Eliminar"
     Exit Sub

End If

Dim i As VbMsgBoxResult

Dim row As Long

row = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0) + 1

i = MsgBox("¿Quieres eliminar el renglón seleccionado?", vbYesNo + vbQuestion, "Eliminar")

If i = vbNo Then Exit Sub

ThisWorkbook.Sheets("Datos").Rows(row).Delete

Call Reset ' Refresca las entradas

MsgBox "El renglón seleccionado se ha eliminado correctamente.", vbOKOnly + vbInformation, "Borrar"
End Sub

Private Sub cmdEdit_Click()
If Selected_List = 0 Then

     MsgBox "No se ha seleccionado ningun renglón.", vbOKOnly + vbInformation, "Editar"
     Exit Sub

End If

Me.txtRowNumber = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0) + 1

'Asignando el registro al control de forms

formulario.txtPaterno.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 1)
formulario.txtMaterno.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 2)
formulario.txtNombre.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 3)
formulario.txtControl.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 4)
formulario.cmbSucursal.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 5)
formulario.txtPuesto.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 6)
formulario.txtDia.value = Format(Me.lstDatabase.List(Me.lstDatabase.ListIndex, 7), "dd/mmm/yyyy")
formulario.txtCaja.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 8)
formulario.txtInventario.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 9)
formulario.txtSobrante.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 10)
formulario.txtObservaciones.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 11)

Me.cmdSubmit.Caption = "Actualizar"
Me.cmdSubmit.BackColor = RGB(255, 165, 0)

MsgBox "Por favor haga los cambios necesarios y haga clic en Actualizar."
End Sub

Private Sub cmdFecha_Click()
Dim dateVariable As Date
dateVariable = CalendarForm.GetDate
formulario.txtDia.value = Format(dateVariable, "dd/mmm/yyyy")
End Sub

Private Sub cmdReset_Click()
Dim i As VbMsgBoxResult

i = MsgBox("¿Quieres limpiar los datos a ingresar?", vbYesNo + vbQuestion, "Reiniciar")

If i = vbNo Then Exit Sub

Call Reset_Form
End Sub

Private Sub cmdSubmit_Click()
'Dim i As VbMsgBoxResult

'i = MsgBox("Quieres agregar la información?", vbYesNo + vbQuestion, "Agregar información")

'If i = vbNo Then Exit Sub

If ValidEntry Then

    Call Submit_Data

End If
End Sub

Private Sub cmdSalir_Click()
'  ActiveWorkbook.Save
  Unload Me ' Cierra el formulario
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub lstDatabase_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If Selected_List = 0 Then
     MsgBox "No hay renglón seleccionado.", vbOKOnly + vbInformation, "Editar"
     Exit Sub
End If

'Me.txtRowNumber = Selected_List + 1 ' Assigning Selected Row Number of Database Sheet

Me.txtRowNumber = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0) + 1

'Asignando el registro al control de forms

formulario.txtPaterno.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 1)
formulario.txtMaterno.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 2)
formulario.txtNombre.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 3)
formulario.txtControl.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 4)
formulario.cmbSucursal.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 5)
formulario.txtPuesto.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 6)
formulario.txtDia.value = Format(Me.lstDatabase.List(Me.lstDatabase.ListIndex, 7), "dd/mmm/yyyy")
formulario.txtCaja.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 8)
formulario.txtInventario.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 9)
formulario.txtSobrante.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 10)
formulario.txtObservaciones.value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 11)

Me.cmdSubmit.Caption = "Actualizar"
Me.cmdSubmit.BackColor = RGB(255, 165, 0)
MsgBox "Porfavor realice los cambio necesarios y haga clic en actualizar."
End Sub

Sub cmdEnviar_Click()
Application.ScreenUpdating = False

Dim App As New Excel.Application
Dim wBook As Excel.Workbook

Dim FileName As String

Dim iRow As Long

FileName = ThisWorkbook.Path & "\concentrado\basededatos.xlsm"

'Revisando si existe el archivo

If Dir(FileName) = "" Then

    MsgBox "No se encuentra el archivo de Base de Datos. Incapaz de proceder.", vbOKOnly + vbCritical, "Error"
    Exit Sub

End If

Set wBook = App.Workbooks.Open(FileName)

App.Visible = False

If wBook.ReadOnly = True Then
    MsgBox "La base de datos esta en uso. Espere un poco y reintente.", vbOKOnly + vbCritical, "Database Busy"
    Exit Sub
End If

'Enviando la informacion

With wBook.Sheets("datos")

    iRow = .Range("A" & Application.Rows.Count).End(xlUp).row + 1

    .Range("A" & iRow).value = iRow - 1
    .Range("B" & iRow).value = formulario.txtPaterno.value   'Apellido paterno
    .Range("C" & iRow).value = formulario.txtMaterno.value   'Apellido materno
    .Range("D" & iRow).value = formulario.txtNombre.value   'Nombre
    .Range("E" & iRow).value = formulario.txtControl.value   'Numero de control
    .Range("F" & iRow).value = formulario.cmbSucursal.value   'Sucursal
    .Range("G" & iRow).value = formulario.txtPuesto.value   'Puesto
    .Range("H" & iRow).value = formulario.txtDia.value   'Correo
    .Range("I" & iRow).value = formulario.txtCaja.value   'Faltante de caja
    .Range("J" & iRow).value = formulario.txtInventario.value   'Faltante de inventario
    .Range("K" & iRow).value = formulario.txtSobrante.value   'Sobrante
    .Range("L" & iRow).value = formulario.txtObservaciones.value   'Observaciones
    .Range("M" & iRow).value = Application.UserName   'Registrado por
    .Range("N" & iRow).value = Format([Now()], "dd/mmm/yyyy HH:MM:SS")  'Fecha de registro

End With

wBook.Close Savechanges:=True

App.Quit

Set App = Nothing

'Reiniciando el formulario

Call Reset_Form

Application.ScreenUpdating = True

MsgBox "¡Información enviada correctamente!"
End Sub

Private Sub txtDia_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim dateVariable As Date
dateVariable = CalendarForm.GetDate
formulario.txtDia.value = Format(dateVariable, "dd/mmm/yyyy")
End Sub

Sub UserForm_Initialize()
    Call Reset_Form
End Sub


