VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Edit Data di ListBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z
Private m_bEditing As Boolean
Private m_lngCurrIndex As Long

Private Sub Command1_Click()
  If Not m_bEditing Then Editing = True
End Sub

Private Sub Form_Load()
  Me.ScaleMode = 3
  Text1.Visible = False
  Text1.Appearance = 0
  Command1.Caption = "Tekan F2 utk mengedit"
  Dim a%
  For a% = 1 To 10
     List1.AddItem "Item yang ke-" & a%
  Next a%
  Set Text1.Font = List1.Font
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, _
Shift As Integer)
  If ((KeyCode = vbKeyF2) And (Shift = 0)) Then
     If (Not m_bEditing) Then Editing = True
  End If
End Sub

Private Sub Text1_LostFocus()
  'Jika textbox kehilangan fokus ketika kita mengedit
  'data, kembalikan data/teks semula dan batalkan
  'proses pengeditan yg telah berlangsung.
  If m_bEditing = True Then
    List1.List(m_lngCurrIndex) = Text1.Tag
    Editing = False
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim strText As String
  If KeyAscii = 10 Or KeyAscii = 13 Then
    If Len(Trim$(Text1.Text)) = 0 Then
       List1.List(m_lngCurrIndex) = Text1.Tag
    Else
       strText = Text1.Text
       'Assginment-kan teks baru ke item data di
       'Listbox ybt
       List1.List(m_lngCurrIndex) = strText
    End If
    Editing = False 'Kembalikan ke posisi semula
    KeyAscii = 0 'Menghindari bunyi beep
  ElseIf KeyAscii = 27 Then 'Jika ditekan Esc untuk
    'membatalkan pengeditan
    List1.List(m_lngCurrIndex) = Text1.Tag 'Kembalikan
    'data semula
    Editing = False
    KeyAscii = 0 'Menghindari bunyi beep
  End If
End Sub

Private Sub Text1_GotFocus()
  'Jika Text1 mendapat fokus, sorot semua isinya.
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_Change()
Dim lpSize As SIZE
Dim phDC As Long
  'Atur ukuran textbox tergantung dari hasil
  'perhitungan ukuran dari textbox dalam pixels
  'Catatan bahwa tingkat perhitungan gagal (untuk
  'beberapa alasan) ketika huruf melebihi dari 14
  'points, tapi jika Anda mempunyai sebuah listbox
  'dengan huruf 14 point, Anda harus men-design-nya
  'dari sana.
  phDC = GetDC(Text1.hwnd)
  If GetTextExtentPoint32(phDC, Text1.Text, _
     Len(Text1.Text), lpSize) = 1 Then
     Text1.Width = Max(50, lpSize.cx)
  End If
  Call ReleaseDC(Text1.hwnd, phDC)
End Sub

Private Property Let Editing(vData As Boolean)
Dim rcItem As RECT
Dim strText As String
Dim lpSize As SIZE
Dim phDC As Long
On Error Resume Next
  'Ambil index dari item data
  m_lngCurrIndex = List1.ListIndex
  '... perlakuan khusus jika tidak ada index
  If m_lngCurrIndex = -1 Then Beep: Exit Property
  
  'Mulai mengedit data...
  If vData = True Then
    strText = List1.List(m_lngCurrIndex)
    If Len(strText) = 0 Then Beep: Exit Property
    'Coba mengambil type RECT dari item dalam list
    If SendMessage(List1.hwnd, LB_GETITEMRECT, _
       ByVal m_lngCurrIndex, rcItem) <> LB_ERR Then
     'Atur RECT. Catatan bahwa ini adalah koordinat di
     'layar Itulah mengapa RECT berhubungan dengan luas
     'dari jendela Listbox. Kita juga mempertimbangkan
     'dengan batas 3-D listbox, jadi jangan memanggil
     'fungsi GetSystemMetrics() jika property Appearence listbox = "Flat"
     With rcItem
      .Left = .Left + List1.Left + _
               GetSystemMetrics(SM_CXEDGE)
      .Top = List1.Top + .Top
      'Mengapa tidak memanggil fungsi GetSysMetrics dan
      'SM_CYEDGE?
      'karena kita ingin data berada di tengah textbox
      'Ambil DC dari listbox lalu hitung tinggi dan
      'lebarnya
      'Catatan bahwa hasil perhitungan gagal (untuk
      'beberapa alasan) ketika ukuran huruf melebihi
      'dari 14 points.
      phDC = GetDC(Text1.hwnd)
      Call GetTextExtentPoint32(phDC, strText, _
           Len(strText), lpSize)
      Call ReleaseDC(Text1.hwnd, phDC)
      'Posisikan dan tampilkan textbox, bawa ke
      'tampilan/urutan teratas.
      Call SetWindowPos(Text1.hwnd, HWND_TOP, .Left, _
           .Top, Max(50, lpSize.cx), _
           lpSize.cy + 2, SWP_SHOWWINDOW Or _
           SWP_NOREDRAW)
     End With
     'Setting property Listbox menyebabkan banyak efek
     'pemunculan, jadi matikan property "redrawing"
     Call SendMessage(List1.hwnd, WM_SETREDRAW, 0, _
      ByVal 0&)
     List1.List(m_lngCurrIndex) = ""
     'Simpan item data dan set fokus ke textbox
     With Text1
       .Enabled = True
       .Tag = strText
       .Text = strText
       .SetFocus
     End With
    End If
  Else
    'Set tanda redraw sehingga listbox menyesuaikan
    'sendiri
    Call SendMessage(List1.hwnd, WM_SETREDRAW, 1, _
         ByVal 0&)
    'Bersihkan isi textbox
    With Text1
      .Enabled = False
      .Visible = False
      .Move 800, 800
      .Text = ""
      .Tag = ""
    End With
    m_lngCurrIndex = -1
  End If
  'Simpan posisi terbaru..........
  m_bEditing = vData
End Property


