Attribute VB_Name = "mdDataEntry"
Sub Reset_Form()

    With frmDataEntr
    
        .txtName.Text = ""
        .txtName.BackColor = vbWhite
        
        .txtDOB.Text = ""
        .txtDOB.BackColor = vbWhite
        
        .optFemale.Value = False
        .optMale.Value = False
        
        .txtMobile.Value = ""
        .txtMobile.BackColor = vbWhite
        
        .txtEmail.Value = ""
        .txtEmail.BackColor = vbWhite
        
        .txtAddress.Value = ""
        .txtAddress.BackColor = vbWhite
        
        .cmbQualification.Clear
        .cmbQualification.BackColor = vbWhite
        
        .cmbQualification.AddItem "10+2"
        .cmbQualification.AddItem "Bachelor Degree"
        .cmbQualification.AddItem "Master Degree"
        .cmbQualification.AddItem "PHD"
        
        .cmbQualification.Value = ""
        
    End With


End Sub
Function ValidEmail(email As String) As Boolean
 Dim oRegEx As Object
 Set oRegEx = CreateObject("VBScript.RegExp")
 With oRegEx
    .Pattern = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
    ValidEmail = .Test(email)
 End With
 Set oRegEx = Nothing
End Function
Function ValidEntry() As Boolean
ValidEntry = True

With frmDataEntr

    'Color predeterminado
    .txtName.BackColor = vbWhite
    .txtDOB.Text = vbWhite
    .txtMobile.BackColor = vbWhite
    .txtEmail.BackColor = vbWhite
    .txtAddress.BackColor = vbWhite
    .cmbQualification.BackColor = vbWhite

    'Validando nombre

    If Trim(.txtName.Value) = "" Then
        MsgBox "El nombre esta en blanco. inserte un nombre valido.", vbOKOnly + vbInformation, "Invalid Entry"
        .txtName.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If

    'Validando fecha de nacimiento

    If Trim(.txtDOB.Value) = "" Then
        MsgBox "Fecha de nacimiento en blanco. Por favor introduzca fecha de nacimiento.", vbOKOnly + vbInformation, "Invalid Entry"
        .txtDOB.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If


    'Validando genero

    If .optFemale.Value = False And .optMale.Value = False Then
        MsgBox "Porfavor, seleccione un genero.", vbOKOnly + vbInformation, "Invalid Entry"
        ValidEntry = False
        Exit Function
    End If

    'Validando grado

    If Trim(.cmbQualification.Value) = "" Then
        MsgBox "Pr favor selecione Grado del menu desplegable.", vbOKOnly + vbInformation, "Invalid Entry"
        .cmbQualification.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If

    'Validando numero de telefono

    If Trim(.txtMobile.Value) = "" Or Len(.txtMobile.Value) < 10 Or Not IsNumeric(.txtMobile.Value) Then
        MsgBox "Por favor introduzca un numero de telefono valido.", vbOKOnly + vbInformation, "Invalid Entry"
        .txtMobile.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If

    'Validando correo electronico

    If ValidEmail(Trim(.txtEmail.Value)) = False Then
        MsgBox "Ingrese un correo electronico valido.", vbOKOnly + vbInformation, "Invalid Entry"
        .txtEmail.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If

    'Validand direccion

    If Trim(.txtAddress.Value) = "" Then
        MsgBox "La direccion esta vacia. introduzca una dirección.", vbOKOnly + vbInformation, "Invalid Entry"
        .txtAddress.BackColor = vbRed
        ValidEntry = False
        Exit Function
    End If

End With
End Function
Sub Submit_Data()
Application.ScreenUpdating = False

Dim App As New Excel.Application
Dim wBook As Excel.Workbook

Dim FileName As String

Dim iRow As Long

FileName = ThisWorkbook.Path & "\basededatos.xlsm"

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

With wBook.Sheets("Database")

    iRow = .Range("A" & Application.Rows.Count).End(xlUp).Row + 1

    .Range("A" & iRow).Value = iRow - 1

    .Range("B" & iRow).Value = frmDataEntr.txtName.Value   'Name

    .Range("C" & iRow).Value = frmDataEntr.txtDOB.Value   'DOB

    .Range("D" & iRow).Value = IIf(frmDataEntr.optFemale.Value = True, "Female", "Male") 'Gender

    .Range("E" & iRow).Value = frmDataEntr.cmbQualification.Value   'Qualification

    .Range("F" & iRow).Value = frmDataEntr.txtMobile.Value   'Mobile Number

    .Range("G" & iRow).Value = frmDataEntr.txtEmail.Value   'Email

    .Range("H" & iRow).Value = frmDataEntr.txtAddress.Value   'Address

    .Range("I" & iRow).Value = Application.UserName   'Submitted By

    .Range("J" & iRow).Value = Format([Now()], "DD-MMM-YYYY HH:MM:SS")  'Submitted On


End With

wBook.Close Savechanges:=True

App.Quit

Set App = Nothing

'Reiniciando el formulario

Call Reset_Form

Application.ScreenUpdating = True

MsgBox "Información actualizada correctamente!"
End Sub
Sub Show_Form()
    frmDataEntr.Show
End Sub



