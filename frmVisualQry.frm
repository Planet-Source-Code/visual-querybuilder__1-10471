VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVisualQry 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7464
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   11412
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7464
   ScaleWidth      =   11412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   252
      Left            =   240
      TabIndex        =   15
      Top             =   4860
      Visible         =   0   'False
      Width           =   11172
      _ExtentX        =   19706
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Max             =   11484
      Orientation     =   1179649
   End
   Begin VB.TextBox txtCause 
      Height          =   288
      Left            =   480
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2292
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   11172
      Begin VB.CommandButton cmdTest 
         Height          =   612
         Left            =   7560
         Picture         =   "frmVisualQry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Test zoekresultaat"
         Top             =   1644
         Width           =   612
      End
      Begin VB.ComboBox cboOrder 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         ItemData        =   "frmVisualQry.frx":08CA
         Left            =   3240
         List            =   "frmVisualQry.frx":08D7
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.CommandButton cmdRestore 
         Height          =   612
         HelpContextID   =   58
         Left            =   8280
         Picture         =   "frmVisualQry.frx":08F8
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Herstellen"
         Top             =   1644
         Width           =   612
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   612
         HelpContextID   =   63
         Left            =   10440
         Picture         =   "frmVisualQry.frx":0A1B
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Afsluiten zonder opslaan"
         Top             =   1644
         Width           =   612
      End
      Begin MSFlexGridLib.MSFlexGrid grdFields 
         Height          =   1212
         Left            =   3180
         TabIndex        =   5
         Top             =   120
         Width           =   7812
         _ExtentX        =   13780
         _ExtentY        =   2138
         _Version        =   393216
         Rows            =   5
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16744576
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox lstTables 
         Height          =   1968
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2172
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sort:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   132
         Left            =   2400
         TabIndex        =   12
         Top             =   870
         Width           =   732
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Visable:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   2400
         TabIndex        =   11
         Top             =   684
         Width           =   732
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Criteria:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   132
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   732
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Table:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   132
         Left            =   2520
         TabIndex        =   9
         Top             =   300
         Width           =   612
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Field:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   132
         Left            =   2520
         TabIndex        =   8
         Top             =   120
         Width           =   612
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tables:"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   20
         Width           =   972
      End
   End
   Begin VB.ListBox List 
      Height          =   2208
      Index           =   0
      ItemData        =   "frmVisualQry.frx":0D25
      Left            =   516
      List            =   "frmVisualQry.frx":0D27
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FF0000&
      Caption         =   "Test"
      ForeColor       =   &H8000000E&
      Height          =   2440
      Index           =   0
      Left            =   -1660
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   240
      X2              =   11400
      Y1              =   5136
      Y2              =   5136
   End
   Begin VB.Image imgJoinRight 
      Height          =   252
      Left            =   840
      Picture         =   "frmVisualQry.frx":0D29
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image imgNorm 
      Height          =   192
      Left            =   480
      Picture         =   "frmVisualQry.frx":116B
      Top             =   3360
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgJoinLeft 
      Height          =   252
      Left            =   840
      Picture         =   "frmVisualQry.frx":12B5
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image2 
      Height          =   252
      Index           =   0
      Left            =   480
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image1 
      Height          =   252
      Index           =   0
      Left            =   480
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image imgDrag 
      Height          =   252
      Left            =   480
      Top             =   3600
      Visible         =   0   'False
      Width           =   1692
   End
End
Attribute VB_Name = "frmVisualQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnnDB As New connection
Dim connectStr As String, DragItem As String
Dim TableCnt As Integer, CurDragNumb As Integer, JoinNumb As Integer, itemCount As Integer
Dim oldX As Single, oldY As Single
Dim imgTop1() As Integer, imgTop2() As Integer, currentscroll As Integer
Dim JoinExp() As String, qryLok As String, JoinFox1() As String, JoinFox2() As String, GridTag() As String

Private Sub SetNew(Itemname As String)
    TableCnt = TableCnt + 1
    Load Frame(TableCnt)
    With Frame(TableCnt)
        .Left = Frame(TableCnt - 1).Left + 1.2 * Frame(TableCnt - 1).Width
        If .Left + .Width > Me.Width Then FlatScrollBar1.Visible = True
        .Caption = Itemname
        .Visible = True
    End With
    Load List(TableCnt)
    With List(TableCnt)
        .Left = Frame(TableCnt).Left
        .Top = Frame(TableCnt).Top + 240
        .Width = 1692
        .Height = 2208
        .ZOrder
        .Visible = True
    End With
    FillItems (Itemname)
End Sub

Public Function VisuallyEditConnection() As Boolean
Dim IUDL As New MSDASC.DataLinks
Dim bResult  As String
    On Error Resume Next
    bResult = IUDL.PromptNew
    If bResult <> "" Then
        connectStr = bResult
        VisuallyEditConnection = True
    End If
End Function

Private Sub SetHeights()
    If List(TableCnt).ListCount < 11 Then
        Frame(TableCnt).Height = List(TableCnt).Height / 10 * List(TableCnt).ListCount + 240
        List(TableCnt).Height = List(TableCnt).Height / 10 * List(TableCnt).ListCount + 240
    End If
End Sub

Private Sub FillItems(Itemname As String)
    Dim fld As Field
    Dim rsDB As Recordset
    Set rsDB = New Recordset
    rsDB.Open "select * from " + Brkts(Itemname) + "", cnnDB, adOpenStatic, adLockReadOnly
    For Each fld In rsDB.Fields
        List(TableCnt).AddItem fld.Name
    Next
    rsDB.Close
    SetHeights
End Sub

Function Brkts(rObjname As Variant) As Variant
  If InStr(rObjname, " ") > 0 And Mid(rObjname, 1, 1) <> "[" Then Brkts = "[" & rObjname & "]" Else Brkts = rObjname
End Function

Private Sub cboOrder_Click()
    If cboOrder.Text <> "Ongesorteerd" Then grdFields.TextArray(cboOrder.Tag) = cboOrder.Text Else grdFields.TextArray(cboOrder.Tag) = ""
    cboOrder.Visible = False
End Sub

Private Sub ComposeQry()
    Dim cnt As Integer
    Dim whereStr As String, OrderStr As String
    If JoinNumb > 0 Then
        qryLok = "Select " & GetFields & " from " & GetHooks
        For cnt = 1 To JoinNumb
            qryLok = qryLok & " " & JoinExp(cnt) & ")"
        Next
        qryLok = Mid(qryLok, 1, Len(qryLok) - 1)
    Else
        qryLok = "Select " & GetFields & " from " & GetTables
    End If
    For cnt = 0 To grdFields.Cols - 1
        If grdFields.TextMatrix(2, cnt) <> "" Then whereStr = whereStr & " and " & grdFields.TextMatrix(2, cnt)
    Next
    If Len(whereStr) > 5 Then whereStr = " where " & Mid(whereStr, 5) Else whereStr = ""
    For cnt = 0 To grdFields.Cols - 1
        If grdFields.TextMatrix(4, cnt) <> "" Then
            Select Case grdFields.TextMatrix(4, cnt)
            Case "Ascending"
                OrderStr = OrderStr & ", " & Brkts(grdFields.TextMatrix(1, cnt)) & "." & Brkts(grdFields.TextMatrix(0, cnt)) & " asc "
            Case "Descending"
                OrderStr = OrderStr & ", " & Brkts(grdFields.TextMatrix(1, cnt)) & "." & Brkts(grdFields.TextMatrix(0, cnt)) & " desc "
            End Select
        End If
    Next
    If Len(OrderStr) > 2 Then OrderStr = " order by " & Mid(OrderStr, 3) Else OrderStr = ""
    qryLok = qryLok & whereStr & OrderStr
End Sub

Private Function GetTables()
    Dim m As Integer
    For m = 1 To TableCnt
        GetTables = GetTables & ", " & Frame(m).Caption
    Next
    GetTables = Mid(GetTables, 2)
End Function

Private Function GetHooks() As String
    Dim m As Integer
    For m = 1 To JoinNumb
        GetHooks = GetHooks & "("
    Next
    GetHooks = Mid(GetHooks, 2)
End Function

Private Function GetFields() As String
    Dim k As Integer, lstcnt As Integer, m As Integer
    Dim trimvar As String
    For k = 1 To TableCnt
        lstcnt = List(k).ListCount
        trimvar = Brkts(Frame(k).Caption)
        For m = 0 To lstcnt - 1
            If List(k).Selected(m) = True Then GetFields = GetFields & trimvar & "." & Brkts(List(k).List(m)) & ","
        Next
    Next
    If Len(GetFields) > 0 Then GetFields = Mid(GetFields, 1, Len(GetFields) - 1) Else GetFields = Brkts(Frame(1).Caption) & ".*"
End Function

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub ResetForm()
    Dim cnt As Integer
    For cnt = 1 To TableCnt
        Unload List(cnt)
        Unload Frame(cnt)
    Next
    TableCnt = 0
    For cnt = 1 To JoinNumb
        Unload Image1(cnt)
        Unload Image2(cnt)
    Next
    FlatScrollBar1.Visible = False
    grdFields.Clear
    itemCount = 0
    JoinNumb = 0
    SetTab
    Me.Refresh
End Sub

Private Sub cmdRestore_Click()
    ResetForm
End Sub

Private Sub cmdTest_Click()
    Dim rsDB As Recordset
    Set rsDB = New Recordset
    Dim cnt As Long, TableCnt As Long
    Dim trimvar As String
    ComposeQry
    On Error GoTo errorhandler
    rsDB.Open qryLok, cnnDB, adOpenStatic
    cnt = rsDB.RecordCount
    rsDB.Close
errorhandler:
    If Err.Number > 0 Then
        MsgBox Err.Description
    Else
        MsgBox "Qrydef:" & vbLf & qryLok & vbLf & vbLf & "Returns:" & vbLf & cnt & " records"
    End If
End Sub

Private Sub FlatScrollBar1_Scroll()
    Dim CurrentControl As Control
    Dim m As Integer
    For Each CurrentControl In Controls
        If CurrentControl.Name = "Frame" Or CurrentControl.Name = "List" Then
            If FlatScrollBar1.Value > currentscroll Then CurrentControl.Left = CurrentControl.Left - FlatScrollBar1.Value + currentscroll Else CurrentControl.Left = CurrentControl.Left + currentscroll - FlatScrollBar1.Value
        End If
    Next
    For m = 1 To JoinNumb
        If FlatScrollBar1.Value > currentscroll Then Image1(m).Left = Image1(m).Left - FlatScrollBar1.Value + currentscroll Else Image1(m).Left = Image1(m).Left + currentscroll - FlatScrollBar1.Value
        If FlatScrollBar1.Value > currentscroll Then Image2(m).Left = Image2(m).Left - FlatScrollBar1.Value + currentscroll Else Image2(m).Left = Image2(m).Left + currentscroll - FlatScrollBar1.Value
    Next
    picLogo.ZOrder
    DrawLines
    currentscroll = FlatScrollBar1.Value
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    If DragItem = "Frame" Then
        If X + Frame(CurDragNumb).Width > Me.Width Then FlatScrollBar1.Visible = True
        Frame(CurDragNumb).Top = Y
        Frame(CurDragNumb).Left = X - oldX
        List(CurDragNumb).Left = Frame(CurDragNumb).Left
        List(CurDragNumb).Top = Frame(CurDragNumb).Top + 240
        If JoinNumb > 0 Then Drawjoins X
    End If
End Sub

Private Sub Drawjoins(X As Single)
    Dim m As Integer
    For m = 1 To JoinNumb
        If JoinFox1(m) = Brkts(Frame(CurDragNumb).Caption) Then
            Image2(m).Left = Frame(CurDragNumb).Left - 250
            Image2(m).Top = Frame(CurDragNumb).Top + (oldY - Image2(m).Top)
        End If
        If JoinFox2(m) = Brkts(Frame(CurDragNumb).Caption) Then
            Image1(m).Left = Frame(CurDragNumb).Left - 250
            Image1(m).Top = Frame(CurDragNumb).Top + (oldY - Image1(m).Top)
        End If
    Next
    DrawLines
End Sub

Private Sub Form_Load()
    Dim m As Integer
    If Not VisuallyEditConnection Then End
    SetDb
    SetTab
    grdFields.Cols = 5
    For m = 1 To 5
        grdFields.ColWidth(m - 1) = 1550
    Next
End Sub

Public Sub SetDb()
    On Error GoTo errorhandler
    If cnnDB.State = adStateOpen Then cnnDB.Close
    cnnDB.CursorLocation = adUseClient
    cnnDB.Open connectStr
errorhandler:
    Exit Sub
End Sub

Private Sub SetTab()
    Dim trimvar As String
    Dim rsDB As Recordset
    Set rsDB = cnnDB.OpenSchema(adSchemaTables)
    lstTables.Clear
    If rsDB.EOF = False Then
        Do
            trimvar = rsDB.Fields!TABLE_NAME
            If (Left(trimvar, 4) <> "MSys" And Left(trimvar, 2) <> "s_" And rsDB.Fields!TABLE_TYPE = "TABLE") Or (rsDB.Fields!TABLE_TYPE = "VIEW" And Left(trimvar, 1) <> "~" And Left(trimvar, 2) <> "s_") Then
                  lstTables.AddItem trimvar
                  lstTables.ItemData(lstTables.NewIndex) = 0
            End If
            rsDB.MoveNext
        Loop Until rsDB.EOF = True
        lstTables.ListIndex = 0
    End If
    rsDB.Close
End Sub

Private Sub Frame_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurDragNumb = Index
    DragItem = "Frame"
    oldX = X
    oldY = Frame(Index).Top
    Frame(Index).Drag vbBeginDrag
End Sub

Private Sub StringReplace(indexNumb As Integer, repString As String)
    If InStr(1, JoinExp(indexNumb), "inner") > 0 Then
        JoinExp(indexNumb) = Mid(JoinExp(indexNumb), 1, InStr(1, JoinExp(indexNumb), "inner") - 1) & repString & " " & Mid(JoinExp(indexNumb), InStr(1, JoinExp(indexNumb), "inner") + 6)
    ElseIf InStr(1, JoinExp(indexNumb), "left") > 0 Then
        JoinExp(indexNumb) = Mid(JoinExp(indexNumb), 1, InStr(1, JoinExp(indexNumb), "left") - 1) & repString & " " & Mid(JoinExp(indexNumb), InStr(1, JoinExp(indexNumb), "left") + 5)
    Else
        JoinExp(indexNumb) = Mid(JoinExp(indexNumb), 1, InStr(1, JoinExp(indexNumb), "right") - 1) & repString & " " & Mid(JoinExp(indexNumb), InStr(1, JoinExp(indexNumb), "right") + 6)
    End If
End Sub


Private Sub grdFields_Click()
    With grdFields
        If .TextMatrix(0, .col) <> "" Then
            Select Case .row
            Case 2
                'load expression builder
            Case 4
                cboOrder.Top = .CellTop + grdFields.Top
                cboOrder.Left = .CellLeft + grdFields.Left
                cboOrder.Width = .CellWidth
                If .Text <> "" Then cboOrder.Text = .Text Else cboOrder.Text = cboOrder.List(0)
                cboOrder.Tag = CellIndex(.row, .col)
                cboOrder.Visible = True
            End Select
        End If
    End With
End Sub

Private Sub grdFields_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(grdFields.TextMatrix(grdFields.MouseRow, grdFields.MouseCol)) > 20 Then grdFields.ToolTipText = grdFields.TextMatrix(grdFields.MouseRow, grdFields.MouseCol) Else grdFields.ToolTipText = ""
End Sub

Private Sub Image1_Click(Index As Integer)
    If Image1(Index).Picture = imgNorm.Picture Then
        If Left(Image1(Index).Tag, 4) = "left" Then Image1(Index).Picture = imgJoinLeft.Picture Else Image1(Index).Picture = imgJoinRight.Picture
        Image2(Index).Picture = imgNorm.Picture
        StringReplace Index, Mid(Image1(Index).Tag, 5)
    Else
        Image1(Index).Picture = imgNorm.Picture
        Image2(Index).Picture = imgNorm.Picture
        StringReplace Index, "inner"
    End If
End Sub

Private Sub Image2_Click(Index As Integer)
    If Image2(Index).Picture = imgNorm.Picture Then
        If Left(Image2(Index).Tag, 4) = "left" Then Image2(Index).Picture = imgJoinLeft.Picture Else Image2(Index).Picture = imgJoinRight.Picture
        Image1(Index).Picture = imgNorm.Picture
        StringReplace Index, Mid(Image1(Index).Tag, 5)
    Else
        Image2(Index).Picture = imgNorm.Picture
        Image1(Index).Picture = imgNorm.Picture
        StringReplace Index, "inner"
    End If
End Sub

Public Sub DrawLines()
    Dim m As Integer
    Me.Refresh
    For m = 1 To JoinNumb
        If Image2(m).Left > Image1(m).Left Then Me.Line (Image2(m).Left, Image2(m).Top + (Image2(m).Height / 2))-(Image1(m).Left + Image1(m).Width, Image1(m).Top + (Image1(m).Height / 2)) Else Me.Line (Image2(m).Left + Image2(m).Width, Image2(m).Top + (Image2(m).Height / 2))-(Image1(m).Left, Image1(m).Top + (Image1(m).Height / 2))
    Next
End Sub

Private Sub SetImages(X As Single, Y As Single)
    Load Image1(JoinNumb)
    With Image1(JoinNumb)
        .Picture = imgNorm.Picture
        .Left = X
        .Top = Y
        .Visible = True
        .ZOrder
    End With
    Load Image2(JoinNumb)
    With Image2(JoinNumb)
        .Picture = imgNorm.Picture
        .Left = oldX
        .Top = oldY
        .Visible = True
        .ZOrder
    End With
    If Image1(JoinNumb).Left > Image2(JoinNumb).Left Then
        Image1(JoinNumb).Tag = "righ"
        Image2(JoinNumb).Tag = "left"
    Else
        Image1(JoinNumb).Tag = "left"
        Image2(JoinNumb).Tag = "righ"
    End If
End Sub

Private Sub DimVars()
    ReDim Preserve JoinExp(JoinNumb)
    ReDim Preserve JoinFox1(JoinNumb)
    ReDim Preserve JoinFox2(JoinNumb)
    ReDim Preserve imgTop1(JoinNumb)
    ReDim Preserve imgTop2(JoinNumb)
    imgTop2(JoinNumb) = Image2(JoinNumb).Top
    imgTop1(JoinNumb) = Image1(JoinNumb).Top
End Sub

Private Sub SetVars(Index As Integer)
    If JoinNumb = 1 Then
        JoinExp(JoinNumb) = Brkts(Frame(Index).Caption) & " inner join " & Brkts(Frame(CurDragNumb).Caption) & " on " & Brkts(Frame(Index).Caption) & "." & Brkts(List(Index).Text) & " = " & Brkts(Frame(CurDragNumb).Caption) & "." & Brkts(List(CurDragNumb).Text)
        JoinFox1(JoinNumb) = Brkts(Frame(CurDragNumb).Caption)
        JoinFox2(JoinNumb) = Brkts(Frame(Index).Caption)
        Image1(JoinNumb).Tag = Image1(JoinNumb).Tag & "right"
        Image2(JoinNumb).Tag = Image2(JoinNumb).Tag & "left"
        Frame(CurDragNumb).Tag = "done"
        Frame(Index).Tag = "done"
    ElseIf Frame(CurDragNumb).Tag <> "done" Then
        JoinExp(JoinNumb) = " inner join " & Brkts(Frame(CurDragNumb).Caption) & " on " & Brkts(Frame(Index).Caption) & "." & Brkts(List(Index).Text) & " = " & Brkts(Frame(CurDragNumb).Caption) & "." & Brkts(List(CurDragNumb).Text)
        JoinFox1(JoinNumb) = Brkts(Frame(CurDragNumb).Caption)
        JoinFox2(JoinNumb) = Brkts(Frame(Index).Caption)
        Image1(JoinNumb).Tag = Image1(JoinNumb).Tag & "right"
        Image2(JoinNumb).Tag = Image2(JoinNumb).Tag & "left"
        Frame(CurDragNumb).Tag = "done"
    Else
        JoinExp(JoinNumb) = " inner join " & Brkts(Frame(Index).Caption) & " on " & Brkts(Frame(CurDragNumb).Caption) & "." & Brkts(List(CurDragNumb).Text) & " = " & Brkts(Frame(Index).Caption) & "." & Brkts(List(Index).Text)
        JoinFox1(JoinNumb) = Brkts(Frame(CurDragNumb).Caption)
        JoinFox2(JoinNumb) = Brkts(Frame(Index).Caption)
        Image1(JoinNumb).Tag = Image1(JoinNumb).Tag & "left"
        Image2(JoinNumb).Tag = Image2(JoinNumb).Tag & "right"
        Frame(Index).Tag = "done"
    End If
End Sub

Private Sub List_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If DragItem = "imgDrag" And Index <> CurDragNumb Then
        If oldX > X + List(Index).Left Then
            X = List(Index).Left + List(Index).Width
            oldX = oldX - List(CurDragNumb).Width - 252
        Else
            X = List(Index).Left - 252
        End If
        Y = Y + List(Index).Top - 100
        JoinNumb = JoinNumb + 1
        SetImages X, Y
        DrawLines
        DimVars
        SetVars (Index)
    End If
    CurDragNumb = 0
End Sub

Private Sub List_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Dim listit As Integer
    If Index <> CurDragNumb Then
        listit = Fix(((Y + 240) / 210)) - 1 + List(Index).TopIndex
        On Error Resume Next
        List(Index).ListIndex = listit
    End If
End Sub

Private Function CellIndex(row As Integer, col As Integer) As Long
    CellIndex = row * grdFields.Cols + col
End Function

Private Sub List_ItemCheck(Index As Integer, Item As Integer)
    Dim m As Integer
    With grdFields
        If List(Index).Selected(Item) = True Then
            If itemCount > 0 Then
                For m = 0 To itemCount - 1
                    If GridTag(CellIndex(0, m)) = Brkts(Frame(Index).Caption) & "." & Brkts(List(Index).Text) Then
                        .TextMatrix(3, m) = "Yes"
                        Exit Sub
                    End If
                Next
            End If
            If itemCount >= .Cols Then
                .Cols = .Cols + 1
                .ColWidth(.Cols - 1) = 1500
            End If
            .TextMatrix(0, itemCount) = List(Index).Text
            .TextMatrix(1, itemCount) = Frame(Index).Caption
            .TextMatrix(3, itemCount) = "Yes"
            ReDim Preserve GridTag(CellIndex(0, itemCount))
            GridTag(CellIndex(0, itemCount)) = Brkts(Frame(Index).Caption) & "." & Brkts(List(Index).Text)
            itemCount = itemCount + 1
        Else
            For m = 0 To itemCount - 1
                If GridTag(CellIndex(0, m)) = Brkts(Frame(Index).Caption) & "." & Brkts(List(Index).Text) Then
                    .TextMatrix(3, m) = "No"
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub List_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim listit As Integer
    If Index <> CurDragNumb Then
        listit = Fix(((Y + 240) / 210)) - 1 + List(Index).TopIndex
        On Error Resume Next
        List(Index).ListIndex = listit
    End If
End Sub

Private Sub List_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim listit As Integer
    If List(Index).Text <> "" And Button = 2 Then
        listit = Fix(((Y + 240) / 210)) - 1 + List(Index).TopIndex
        On Error Resume Next
        List(Index).ListIndex = listit
        oldX = List(Index).Left + List(Index).Width
        oldY = Y + List(Index).Top
        DragItem = "imgDrag"
        CurDragNumb = Index
        With imgDrag
            .Top = oldY
            .Left = Frame(Index).Left + 200
            .Width = Len(List(Index).Text) * 200
            .Drag vbBeginDrag
            .Visible = True
        End With
    End If
End Sub

Private Sub List_Scroll(Index As Integer)
    Dim m As Integer
    For m = 1 To JoinNumb
        If JoinFox1(m) = Brkts(Frame(Index).Caption) Then
            If imgTop2(m) - (210 * List(Index).TopIndex) > List(Index).Top - 240 Then Image2(m).Top = imgTop2(m) - (210 * (List(Index).TopIndex))
        ElseIf JoinFox2(m) = Brkts(Frame(Index).Caption) Then
            If imgTop1(m) - (210 * List(Index).TopIndex) > List(Index).Top - 240 Then Image1(m).Top = imgTop1(m) - (210 * (List(Index).TopIndex))
        End If
    Next
    DrawLines
End Sub

Private Sub lstTables_DblClick()
    If lstTables.Text <> "" Then
        SetNew lstTables.Text
        lstTables.RemoveItem (lstTables.ListIndex)
    End If
End Sub
