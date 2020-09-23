VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SnifferProTCP"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   2640
      Width           =   8895
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3015
         Left            =   0
         TabIndex        =   5
         Top             =   165
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5318
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3015
         Left            =   3600
         TabIndex        =   4
         Top             =   165
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5318
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   50000
         TextRTF         =   $"Form1.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0946
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1408
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1722
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2630
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2BCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Tag             =   "1"
      Top             =   5880
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   873
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tab             =   2
      TabHeight       =   670
      TabCaption(0)   =   " Packet"
      TabPicture(0)   =   "Form1.frx":34A4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "  Data"
      TabPicture(1)   =   "Form1.frx":3560
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "  DataPacket"
      TabPicture(2)   =   "Form1.frx":360D
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   -71280
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As Integer
Dim es As String

Private Sub Form_Load()
Me.Move (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)

z = 0
ListView1.View = lvwReport
TreeView1.ImageList = ImageList1
ListView1.ColumnHeaders.Add 1, , "IP header", 1500
ListView1.ColumnHeaders.Add 2, , "Sport", 1000
ListView1.ColumnHeaders.Add 3, , "IP Destino", 1500
ListView1.ColumnHeaders.Add 4, , "Dport", 1000
ListView1.ColumnHeaders.Add 5, , "Protocol", 1000
ListView1.ColumnHeaders.Add 6, , "Time", 1500
salir = False
End Sub




Private Sub ListView1_Click()
Dim coun As Long
Dim sar As String, sar3 As String
Dim sar1 As String, sar2 As String

RichTextBox1.Text = ""
Dim buffer() As Byte
buffer = str
For i = 0 To tamaño(ListView1.SelectedItem.Index)
'StrConv(str, vbUnicode)
coun = coun + 1
If Len(Hex(buffer(i))) = 1 Then
sar = "0" & Hex(buffer(i))
Else
sar = Hex(buffer(i))
End If
sar3 = sar3 & sar

If Asc(Chr("&h" & Hex(buffer(i)))) < 32 Then
sar1 = "."
Else
sar1 = Chr("&h" & Hex(buffer(i)))
End If
sar2 = sar2 & sar1
RichTextBox1.Text = RichTextBox1.Text & sar & " "

If coun = 15 Then

RichTextBox1.Text = RichTextBox1.Text & " |" & sar2 & vbCrLf: coun = 0: sar2 = "": sar3 = ""



End If




Next i

If coun < 15 Then

r = 44 - (coun * 3) + 1

es = String(r, Chr(32))
RichTextBox1.Text = RichTextBox1.Text & es & " |" & sar2

End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
z = z + 1
Recibir s, 1
If salir = True Then WCleanup s: salir = False: MsgBox "Se salió del dumpeo", vbInformation Or vbOKOnly, "SniffProTCP"
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case SSTab1.Tab
Case 0
TreeView1.Visible = True
RichTextBox1.Visible = False
TreeView1.Width = RichTextBox1.Left + RichTextBox1.Width
Case 1
TreeView1.Visible = False
RichTextBox1.Visible = True
RichTextBox1.Left = TreeView1.Left + 10
RichTextBox1.Width = Frame1.Width - 20
Case 2
TreeView1.Visible = True
RichTextBox1.Visible = Visible
TreeView1.Left = 0
TreeView1.Width = 3495
RichTextBox1.Left = 3600
RichTextBox1.Width = 5175
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
TreeView1.Nodes.Clear
RichTextBox1.Text = ""
ListView1.ListItems.Clear
Case 3
TreeView1.Nodes.Clear
RichTextBox1.Text = ""
ListView1.ListItems.Clear
Connecting ip(hostname), Picture1
Case 4
salir = True
Case 6
Form2.Show vbModal
Case 7
Shell "explorer.exe " & App.Path & "\help\Untitled-1.htm", vbNormalFocus
End Select
End Sub
