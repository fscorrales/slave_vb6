VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEditarFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variable de objetos con eventos para el FlexGrid y el Textbox

Private WithEvents UnFlexGrid As MSHFlexGrid
Attribute UnFlexGrid.VB_VarHelpID = -1
Private WithEvents UnText As MaskEdBox
Attribute UnText.VB_VarHelpID = -1
Private strEstructuraACambiar As String


'Inicia la clase (Recibe el FlexGrid y el TextBox asociado)
Public Sub Iniciar(ElFlexGrid As MSHFlexGrid, ElTextBox As MaskEdBox)

Set UnFlexGrid = ElFlexGrid
Set UnText = ElTextBox

End Sub

'Al hacer click en el FlexGrid se hace visible, se le da foco _
y se ajustan otros propiedades
Private Sub UnFlexGrid_Click()
  
  If ValidarEstructuraPresupuestaria(UnFlexGrid.Text) = False Then
    AjustarText
    strEstructuraACambiar = UnFlexGrid.Text
    UnText.Visible = True
    UnText.SetFocus
  End If
  
End Sub

Private Sub UnFlexGrid_GotFocus()

    UnFlexGrid_RowColChange

End Sub

Private Sub UnFlexGrid_LeaveCell()
    'If UnText.Text <> "  -  -  -   " Then
    If ValidarEstructuraPresupuestaria(UnText.Text) = True _
    And UnText.Text <> strEstructuraACambiar _
    And strEstructuraACambiar <> "" Then
        UnFlexGrid.Text = UnText.Text
        'Procedimos a editar la estructura presupuestaria
        SQL = "Select * from PRECARIZADOS Where AGENTES = " & "'" & UnFlexGrid.TextMatrix(UnFlexGrid.Row, 0) & "'"
        Set rstRegistroSlave = New ADODB.Recordset
        rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        SQL = ""
        rstRegistroSlave!AGENTES = UnFlexGrid.TextMatrix(UnFlexGrid.Row, 0)
        rstRegistroSlave!ACTIVIDAD = Left(UnText.Text, 8)
        rstRegistroSlave!PARTIDA = Right(UnText.Text, 3)
        rstRegistroSlave.Update
        rstRegistroSlave.Close
        Set rstRegistroSlave = New ADODB.Recordset
        UnFlexGrid.CellBackColor = &H80000005
    End If
    
End Sub

Private Sub UnFlexGrid_RowColChange()
  Static oRow As Integer, oCol As Integer, Cambio As Boolean
   
    If Cambio Then Exit Sub
    Cambio = True
     
    With UnFlexGrid
      If .col <> oCol Or .Row <> oRow Then
        oRow = .Row
        oCol = .col
        
        AjustarText
        If ValidarEstructuraPresupuestaria(UnFlexGrid.Text) = False Then
          UnText.Visible = True
          UnText.SetFocus
        End If
        
      End If
    End With
  Cambio = False
End Sub

'Private Sub UnText_KeyDown(KeyCode As Integer, Shift As Integer)
'  With UnFlexGrid
'    Select Case KeyCode
'
'      Case vbKeyRight
'        If .Col + 2 > .Cols And .Row + 1 < .Rows Then
'            UnFlexGrid_LeaveCell
'          .Col = 2
'          .Row = .Row + 1
'        ElseIf .Col + 1 < .Cols And .Row < .Rows Then
'            UnFlexGrid_LeaveCell
'         .Col = .Col + 1
'        End If
'
'      Case vbKeyUp
'        If .Row - 1 > 0 Then
'            UnFlexGrid_LeaveCell
'            .Row = .Row - 1
'        End If
'      Case vbKeyDown, vbKeyReturn
'        If .Row + 1 < .Rows Then
'            UnFlexGrid_LeaveCell
'            .Row = .Row + 1
'        End If
'      Case vbKeyLeft
'        If .Col - 2 = 0 And .Row - 1 <> 0 Then
'            UnFlexGrid_LeaveCell
'          .Col = 2
'          .Row = .Row - 1
'        ElseIf .Col - 2 <> 0 Then
'            UnFlexGrid_LeaveCell
'          .Col = .Col - 1
'        End If
'      End Select
'    End With
'    UnFlexGrid_RowColChange
'End Sub

'Si se presiona Enter o el Tab se cambia de celda, es decir se _
ubica el textbox
Private Sub UnText_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then KeyAscii = 0
  If KeyAscii = vbKeyTab Then
    'Call UnText_KeyDown(vbKeyRight, 0)
    With UnFlexGrid
        If .Row + 1 < .Rows Then
            UnFlexGrid_LeaveCell
            .Row = .Row + 1
        End If
    End With
    KeyAscii = 0
  End If
End Sub

Private Sub Untext_LostFocus()
    UnText.Visible = False
    strEstructuraACambiar = ""
    UnFlexGrid.Text = UnText.Text
    UnFlexGrid_RowColChange
End Sub

Private Sub AjustarText()
  With UnFlexGrid
    UnText.Text = .Text
    UnText.FontSize = .Font.Size
    UnText.Height = .CellHeight
    
    If .CellLeft + .CellWidth > .Width Then
        UnText.Width = .Width - .CellLeft
    Else
        UnText.Width = .CellWidth
    End If
    
    UnText.Left = .CellLeft + .Left
    UnText.Top = .CellTop + .Top
  End With
End Sub


