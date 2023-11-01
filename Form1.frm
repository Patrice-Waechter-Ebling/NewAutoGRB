VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rstemploye As ADODB.Recordset
Dim gllaray(1 To 1500, 1 To 2) As Variant
Dim r As Integer
Dim oXLApp As Excel.Application         'Declare the object variables
Dim oXLBook As Excel.Workbook
Dim oXLSheet As Excel.Worksheet
Set oXLApp = New Excel.Application    'Create a new instance of Excel
Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first worksheet
oXLApp.Visible = False


Set rstemploye = New ADODB.Recordset
r = 1
Call rstemploye.Open("SELECT groupe,employe FROM GRB_Employés", g_connData, adOpenDynamic, adLockOptimistic)
Do While Not rstemploye.EOF
gllaray(r, 1) = rstemploye.Fields("groupe")
gllaray(r, 2) = rstemploye.Fields("employe")
r = r + 1
rstemploye.MoveNext
Loop



oXLSheet.Range("A1:B1").Font.Bold = True
oXLSheet.Range("A:b").HorizontalAlignment = xlRight
oXLSheet.Range("A1: B1").Value = Array("Groupe", "Nom") 'GLL



'inscription des valeur du tableau dans excel
oXLSheet.Range("A2").Resize(r, 2).Value = gllaray

'ajustement largeur des colonne
oXLSheet.Range("A:b").Columns.AutoFit
oXLApp.Visible = True

  Call rstemploye.Close
Set rstemploye = Nothing
        


End Sub

