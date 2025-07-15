Attribute VB_Name = "mdEntradaDatos"

Sub Reset_Form()
Dim iRow As Long

With formulario

    .txtPaterno.Text = ""
    .txtPaterno.BackColor = vbWhite
    
    .txtMaterno.Text = ""
    .txtMaterno.BackColor = vbWhite
    
    .txtNombre.Text = ""
    .txtNombre.BackColor = vbWhite

    .txtControl.Text = ""
    .txtControl.BackColor = vbWhite
    
    .txtPuesto.Text = ""
    .txtPuesto.BackColor = vbWhite

    .txtDia.Text = Format([Now()], "dd/mmm/yyyy")
    .txtDia.BackColor = vbWhite

    .txtCaja.value = "0.0"
    .txtCaja.BackColor = vbWhite

    .txtInventario.value = "0.0"
    .txtInventario.BackColor = vbWhite

    .txtSobrante.value = "0.0"
    .txtSobrante.BackColor = vbWhite

    .txtObservaciones.value = "Ninguna"
    .txtObservaciones.BackColor = vbWhite

    .cmdSubmit.Caption = "Agregar"

    '.cmbCourse.Clear
    .cmbSucursal.BackColor = vbWhite

    'Rango dinamico basado en sucursales
    shSucursales.Range("A2", shSucursales.Range("A" & Rows.Count).End(xlUp)).Name = "Dynamic"

    .cmbSucursal.RowSource = "Dynamic"

    .cmbSucursal.value = ""

    'Informacion visible para isDatabase
    
    .lstDatabase.ColumnCount = 14
    .lstDatabase.ColumnHeads = False
    
    .lstDatabase.ColumnWidths = "0;70;70;70;60;65;60;70;70;75;55;80;0;0"

    iRow = shDatos.Range("A" & Rows.Count).End(xlUp).row + 1 ' Identifica la última columna en blanco
    'iRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row
    
    If iRow > 14 Then inicio = iRow - 12 Else inicio = 1

    .lstDatabase.RowSource = "Datos!A" & inicio & ":N" & iRow + 1
    .lstDatabase.ListIndex = -1
    .txtRowNumber = iRow
    .cmdSubmit.BackColor = &H8000000F
End With
End Sub

Function ValidEntry() As Boolean

ValidEntry = True

With formulario

    'Color predeterminado

    .txtPaterno.BackColor = vbWhite
    .txtMaterno.BackColor = vbWhite
    .txtNombre.BackColor = vbWhite
    .txtControl.BackColor = vbWhite
    .cmbSucursal.BackColor = vbWhite
    .txtPuesto.BackColor = vbWhite
    .txtDia.BackColor = vbWhite
    .txtCaja.BackColor = vbWhite
    .txtInventario.BackColor = vbWhite
    .txtObservaciones.BackColor = vbWhite
    
     'validando paterno

    If Trim(.txtPaterno.value) = "" Then
        MsgBox "Introducir apellido paterno correctamente.", vbOKOnly + vbInformation, "Apellido"
        .txtPaterno.BackColor = vbRed
        .txtPaterno.SetFocus
        ValidEntry = False
        Exit Function
    End If
    
    'validando nombre

    If Trim(.txtNombre.value) = "" Then
        MsgBox "Introducir nombre correctamente.", vbOKOnly + vbInformation, "Nombre"
        .txtNombre.BackColor = vbRed
        .txtNombre.SetFocus
        ValidEntry = False
        Exit Function
    End If
    
    'validando numero de control

    If Trim(.txtControl.value) = "" Then
        MsgBox "Introduzca un número de control valido.", vbOKOnly + vbInformation, "Entrada invalida"
        .txtControl.BackColor = vbRed
        .txtControl.SetFocus
        ValidEntry = False
        Exit Function
    End If
    
    'validando sucursal

    If Trim(.cmbSucursal.value) = "" Then
        MsgBox "Seleccione sucursal del menú.", vbOKOnly + vbInformation, "Sucursal"
        .cmbSucursal.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If


    'validando puesto

    If Trim(.txtPuesto.value) = "" Then
        MsgBox "Introducir puesto.", vbOKOnly + vbInformation, "Puesto"
        .txtPuesto.BackColor = vbRed
        .txtPuesto.SetFocus
        ValidEntry = False
        Exit Function
    End If

    'validando fecha de corte

    If Trim(.txtDia.value) = "" Then
        MsgBox "Introduzca fecha de corte.", vbOKOnly + vbInformation, "Entrada invalida"
        .txtDia.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If
    
        'Validando faltante de caja

    If Trim(.txtCaja.value) = "" Or Not IsNumeric(.txtCaja.value) Then
        MsgBox "Por favor indroduzca una cantida valida.", vbOKOnly + vbInformation, "Entrada invalida"
        .txtCaja.BackColor = vbRed
        .txtCaja.SetFocus
        ValidEntry = False
        Exit Function
    End If
    
        'Validando faltante de inventario

    If Trim(.txtInventario.value) = "" Or Not IsNumeric(.txtInventario.value) Then
        MsgBox "Por favor indroduzca una cantida valida.", vbOKOnly + vbInformation, "Entrada invalida"
        .txtInventario.BackColor = vbRed
        .txtInventario.SetFocus
        ValidEntry = False
        Exit Function
    End If

        'Validando sobrante

    If Trim(.txtSobrante.value) = "" Or Not IsNumeric(.txtSobrante.value) Then
        MsgBox "Por favor indroduzca una cantida valida.", vbOKOnly + vbInformation, "Entrada invalida"
        .txtSobrante.BackColor = vbRed
        .txtSobrante.SetFocus
        ValidEntry = False
        Exit Function
    End If

    'validando observaciones

    If Trim(.txtObservaciones.value) = "" Then
        MsgBox "Introduzca observaciones.", vbOKOnly + vbInformation, "Entrada invalida"
        .txtObservaciones.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If

End With
End Function

Sub Submit_Data()
Dim iRow As Long

If formulario.txtRowNumber.value = "" Then
   
 iRow = shDatos.Range("A" & Rows.Count).End(xlUp).row + 1 ' Identifica último renglón en blanco
 'iRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row

Else
    iRow = formulario.txtRowNumber.value
End If

With shDatos.Range("A" & iRow)

.Offset(0, 0).value = "=Row()-1" ' ID
.Offset(0, 1).value = UCase(formulario.txtPaterno.value) 'Apellido paterno
.Offset(0, 2).value = UCase(formulario.txtMaterno.value) 'Apellido materno
.Offset(0, 3).value = UCase(formulario.txtNombre.value) 'Nombre
.Offset(0, 4).value = formulario.txtControl.value 'Numero de control
.Offset(0, 5).value = formulario.cmbSucursal.value    'Sucursal
.Offset(0, 6).value = UCase(formulario.txtPuesto.value) 'Puesto
.Offset(0, 7).value = Format(formulario.txtDia.value, "dd/mmm/yyyy") 'Dia del corte
.Offset(0, 8).value = formulario.txtCaja.value 'Faltante de caja
.Offset(0, 9).value = formulario.txtInventario.value 'Faltante de inventario
.Offset(0, 10).value = formulario.txtSobrante.value 'Sobrante
.Offset(0, 11).value = formulario.txtObservaciones.value 'Observaciones
.Offset(0, 12).value = Application.UserName    'Registrado por
.Offset(0, 13).value = Format([Now()], "dd/mmm/yyyy HH:MM:SS")   'Hora del registro

'Limpia el formulario
End With

Call Reset_Form
Call Reset_Form
Application.ScreenUpdating = True

'MsgBox "Información registrada correctamente"
End Sub

Function Selected_List() As Long
Dim i As Long
Selected_List = 0
'If formulario.lstDatabase.ListCount = 1 Then Exit Function ' If no items exist in List Box
For i = 0 To formulario.lstDatabase.ListCount - 1
If formulario.lstDatabase.Selected(i) = True Then
   Selected_List = i + 1
   Exit For
End If
Next i
End Function


Sub Show_Form()
formulario.Show
End Sub
