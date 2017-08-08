VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form_Entri_Barang 
   BackColor       =   &H0000C000&
   Caption         =   "Entri Data Barang"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15675
   ControlBox      =   0   'False
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Entri "
   ScaleHeight     =   6720
   ScaleWidth      =   15675
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Height          =   1320
      Left            =   9000
      ScaleHeight     =   1260
      ScaleWidth      =   5835
      TabIndex        =   20
      Top             =   4680
      Width           =   5895
   End
   Begin VB.PictureBox Picture2 
      Height          =   3480
      Left            =   8880
      Picture         =   "Form6.frx":628A
      ScaleHeight     =   3481.901
      ScaleMode       =   0  'User
      ScaleWidth      =   6075
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   480
      Picture         =   "Form6.frx":49DCC
      ScaleHeight     =   1155
      ScaleWidth      =   2115
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   3480
      Left            =   8880
      ScaleHeight     =   3420
      ScaleWidth      =   6075
      TabIndex        =   17
      Top             =   960
      Width           =   6135
   End
   Begin VB.CommandButton cmd_Barcode 
      Caption         =   "Buat Barcode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txt_nama_supplier 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   3855
   End
   Begin VB.ComboBox cb_kategori 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin MSComctlLib.ListView list_supplier 
      Height          =   2055
      Left            =   3120
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "KODE"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NAMA SUPLIER"
         Object.Width           =   5999
      EndProperty
   End
   Begin VB.CommandButton btn_kategori 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txt_kode_supplier 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   735
   End
   Begin MSMask.MaskEdBox txt_jual 
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt_nama 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   5175
   End
   Begin VB.TextBox txt_kode 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton btn_cancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btn_save 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C000&
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C000&
      Caption         =   "Harga Jual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "Kategori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Kode Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ENTRI dan UPDATE DATA BARANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "Form_Entri_Barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO1
    bmiHeader As BITMAPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type

Private Type BITMAPINFO8
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

 Private Declare Function CreateDIBSection1 Lib "gdi32" _
    Alias "CreateDIBSection" (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO1, ByVal un As Long, _
    ByVal lplpVoid As Long, ByVal handle As Long, _
    ByVal dw As Long) As Long

Private Declare Function CreateDIBSection8 Lib "gdi32" _
    Alias "CreateDIBSection" (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO8, ByVal un As Long, _
    ByVal lplpVoid As Long, ByVal handle As Long, _
    ByVal dw As Long) As Long

Private Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Dim rsBarang As New ADODB.Recordset
Dim txt_sup_toggle As Boolean

Private Sub btn_kategori_Click()
    Dim new_kategori As String
    new_kategori = InputBox("Kategori Baru: ", "Kategori")
    
    If new_kategori = "" Then
        Exit Sub
    End If
    
    Dim i As Integer
    i = 0
    Do While i < cb_kategori.ListCount
        If Trim(UCase(new_kategori)) = Trim(UCase(cb_kategori.List(i))) Then
            MsgBox "Kategori telah terdaftar"
            Exit Sub
        End If
        i = i + 1
    Loop
    
    cb_kategori.Text = new_kategori
    cb_kategori.AddItem (new_kategori)
    con.Execute ("insert into tbkategori values('" & new_kategori & "')")
End Sub

Private Sub cb_kategori_KeyPress(key As Integer)
    If key = 13 Then
        txt_jual.SetFocus
    End If
End Sub

Private Sub cmd_Barcode_Click()
    AddFont (App.Path & "\fre3of9x.ttf")
    Dim resize As Variant
    Picture1.Width = 6135
    'Picture1.Height = 4337
    Picture1.Height = 3480
    Picture1.Cls
    DoEvents
    resize = 0.5
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.PaintPicture Picture2.Image, 0, 0
    
    Picture4.AutoRedraw = True
    Picture4.AutoSize = True
    Picture4.FontBold = False
    Picture4.FontSize = 72
    Picture4.Font = "Free 3 of 9 Extended"
    Picture4.CurrentY = 10
    Picture4.CurrentX = (Picture4.ScaleWidth - Picture4.TextWidth(Barcode39(txt_kode.Text))) / 2
    Picture4.Print Barcode39(txt_kode.Text)
'    Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(Barcode39(txt_kode.Text))) / 2
'    Picture1.Print Barcode39(txt_kode.Text)
'    Picture4.Font = "dotumche"
    DoEvents
'    Picture1.AutoRedraw = True
'    Picture1.AutoSize = True
'    Picture1.PaintPicture Picture4.Image, 100, 1500, resize * Picture4.Width, resize * Picture4.Height
    '(Picture1.Width - (Picture4.Width * resize)) / 4
    Call MonoChrome
    
    Picture1.FontSize = 24
    Picture1.FontBold = True
    Picture1.Font = "Times New Roman"
    Picture1.FontSize = 14
    Picture1.FontBold = True
    Picture1.CurrentY = 300
    Picture1.CurrentX = Picture1.ScaleWidth - Picture1.TextWidth("Tenant No :   ")
    Picture1.Print "Tenant No :   "
    Picture1.FontSize = 36
    Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(txt_kode_supplier.Text)) * 0.88
    Picture1.Print txt_kode_supplier.Text
    Picture1.FontBold = False
    Picture1.FontSize = 14
'    Picture1.Print Tab(1); "                                          "
'    Picture1.Print Tab(1); "                                          "
'    Picture1.Print Tab(1); "                                          "
'    Picture1.Print Tab(1); "                                          "
'    Picture1.Print Tab(1); "                                          "
    'Picture1.Print Tab(1); "                                          "
    Picture1.FontBold = True
    Picture1.CurrentY = 1200
    Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(txt_nama.Text)) / 2
    'Picture1.Print Tab(((38 - Len(txt_nama.Text)) / 2) + 1); txt_nama.Text
    Picture1.Print txt_nama.Text
    'Picture1.Print Tab(2); Barcode39(txt_kode.Text)
    
'    Picture1.Print Tab(((24 - Len(txt_kode.Text)) / 2) + 1); txt_kode.Text
    Picture1.CurrentY = 2500
    Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(txt_kode.Text)) / 2
    Picture1.Print txt_kode.Text
    Picture1.Print Tab(1); "                                          "
    Picture1.Print Tab(1); "                                          "
    Picture1.FontSize = 14
'    Picture1.Print Tab(37 - Len(txt_kode_supplier.Text)); txt_kode_supplier.Text
    Picture1.FontBold = False

    Picture1.PaintPicture Picture3.Image, 10, 10
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORLandscape

    Dim xMargin, yMargin As Integer
    xMargin = 300
    yMargin = 100
    resize = 0.66
    Dim x, y As Integer
    Dim posx, posy As Variant

    For x = 0 To 0
        posx = x * resize * Picture1.Width + xMargin
        For y = 1 To 1
            posy = y * resize * Picture1.Height + yMargin
            Printer.PaintPicture Picture1.Image, posx, posy, Picture1.Width * resize, Picture1.Height * resize
            Printer.PaintPicture Picture4.Image, posx + (Picture1.Width * resize - Picture4.Width * resize * 0.5) / 2, posy + 1090, Picture4.Width * resize * 0.5, Picture4.Height * resize * 0.5
        Next
    Next
    Printer.EndDoc
    RemoveFont (App.Path & "\fre3of9x.ttf")
    DoEvents
End Sub

Private Sub Form_Activate()
    If txt_kode = "" Then
        txt_kode.SetFocus
    Else
        txt_nama.SetFocus
    End If
End Sub

Private Sub btn_save_Click()
    Dim a As New ADODB.Recordset
    
    'kerjakan cek kategori
    
    If cek_Kategori = False Then
        MsgBox "Kategori tidak ditemukan"
        Exit Sub
    End If
    
    'kerjakan
    
    'If Trim(txt_kode.Text) = "" Or txt_nama.Text = "" Or txt_modal = "" Or txt_jual = "" Or txt_kode_supplier = "" Then
    If Trim(txt_kode.Text) = "" Or txt_nama.Text = "" Or txt_jual = "" Or txt_kode_supplier = "" Then
        MsgBox "Isi Data dengan Lengkap.....!"
        Exit Sub
    End If
    
    If getBarang(txt_kode) Then
        'disabled, hapus jumlah_akhir
        'con.Execute ("Update tbbarang set nama='" & txt_nama & "',kategori='" & cb_kategori.Text & "',harga_modal='" & Val(txt_modal) & "',harga_jual='" & Val(txt_jual) & "',kdsuplier='" & Val(txt_kode_supplier) & "',tgl_masuk='" & Format(dp_masuk, "yyyy-MM-dd") & "',ketahanan='" & Val(txt_ketahanan) & "', jumlah_akhir=" & Val(txt_stok) & " where kode='" & Trim(txt_kode) & "' ")
        'con.Execute ("Update tbbarang set nama='" & txt_nama & "', kategori='" & cb_kategori.Text & "', harga_modal = " & Val(txt_modal) & ", harga_jual = " & Val(txt_jual) & ", kdsuplier='" & Val(txt_kode_supplier) & "' where kode='" & Trim(txt_kode.Text) & "' ")
        con.Execute ("Update tbbarang set nama='" & txt_nama & "', kategori='" & cb_kategori.Text & "', harga_jual = " & Val(txt_jual) & ", kdsuplier='" & Val(txt_kode_supplier) & "' where kode='" & Trim(txt_kode.Text) & "' ")
    Else
        'disabled, hapus jumlah_akhir
        'con.Execute ("Insert into tbbarang values('" & Trim(txt_kode) & "' ,'" & txt_nama & "','" & cb_kategori.Text & "','" & Val(txt_modal) & "','" & Val(txt_jual) & "'," & Val(txt_stok) & ",'" & Val(txt_kode_supplier) & "','" & Format(dp_masuk, "yyyy-MM-dd") & "', '" & Val(txt_ketahanan) & "')")
        'con.Execute ("Insert into tbbarang values('" & Trim(txt_kode) & "' ,'" & txt_nama & "','" & cb_kategori.Text & "' ,'" & Val(txt_modal) & "','" & Val(txt_jual) & "','" & Val(txt_kode_supplier) & "')")
        con.Execute ("Insert into tbbarang values('" & Trim(txt_kode) & "' ,'" & txt_nama & "' ,'" & cb_kategori.Text & "' ,'" & Val(txt_jual) & "' ,'" & Val(txt_kode_supplier) & "')")
    
    End If
    kosongkan
    
    Form_List_barang.refreshlist
    Unload Me
End Sub

Sub kosongkan()
    txt_kode = ""
    txt_nama = ""
    cb_kategori.ListIndex = -1
    txt_kode_supplier = ""
    txt_nama_supplier = ""
    'txt_ketahanan = ""
    'txt_modal = 0
    txt_jual = 0
End Sub
Private Sub btn_cancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Set rsBarang = con.Execute("select * from tbbarang")
    
    kosongkan
    'dp_masuk = Date
    txt_sup_toggle = True
    reload_Kategori
    reload_Supplier
    
    list_supplier.Visible = False
    
    'lbl_stok.Visible = isMaster
    'txt_stok.Visible = isMaster
End Sub

Private Function getBarang(kode As String) As Boolean
    'If rsbarang.EOF Or rsbarang.BOF Then
     '   getBarang = False
      '  Exit Function
    'End If
    
    Dim found As Boolean
    found = False
    rsBarang.MoveFirst
    Do While Not rsBarang.EOF
        If rsBarang!kode = kode Then
            found = True
            Exit Do
        End If
        rsBarang.MoveNext
    Loop
    getBarang = found
End Function

Private Sub txt_kode_change()
    If getBarang(txt_kode) Then
        txt_nama = rsBarang!nama
        cb_kategori.Text = rsBarang!kategori
        'txt_modal.Text = rsbarang!harga_modal
        txt_jual = rsBarang!harga_jual
        txt_kode_supplier.Text = rsBarang!kdsuplier
        Set rsSupplier = con.Execute("select * from tbsuplier")
        If getSupplier(rsBarang!kdsuplier) Then
            txt_nama_supplier = rsSupplier!nmsuplier
        End If
        'txt_ketahanan.Text = rsbarang!ketahanan
        'dp_masuk = rsbarang!tgl_masuk
        'txt_stok = rsbarang!jumlah_akhir
    Else
        txt_nama.Text = ""
        cb_kategori.ListIndex = -1
        'txt_modal = 0
        txt_jual = 0
        txt_kode_supplier = ""
        txt_nama_supplier = ""
        'txt_ketahanan = ""
        'txt_stok = 0
    End If
    txt_sup_toggle = True
End Sub
Private Sub txt_kode_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_nama.SetFocus
    End If
End Sub
Private Sub txt_nama_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cb_kategori.SetFocus
    End If
End Sub

Private Sub txt_modal_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_jual.SetFocus
    End If
End Sub
Private Sub txt_jual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_kode_supplier.SetFocus
    End If
End Sub

Private Sub txt_kode_supplier_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        txt_sup_toggle = True
        
        Set rsSupplier = con.Execute("select * from tbsuplier")
        
        If getSupplier(txt_kode_supplier) Then
            txt_nama_supplier.Text = rsSupplier!nmsuplier
            'txt_ketahanan.SetFocus
            btn_save.SetFocus
        Else
            MsgBox "Supplier tidak terdaftar"
            txt_kode_supplier.Text = ""
        End If
    Else
        txt_nama_supplier = ""
    End If
End Sub

Private Sub txt_nama_supplier_Change()
    If txt_nama_supplier.Text <> "" And txt_sup_toggle = False Then
        list_supplier.Visible = True
        reload_Supplier
    Else
        list_supplier.Visible = False
        txt_sup_toggle = False
    End If
End Sub

'Private Sub txt_ketahanan_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Val(txt_ketahanan) > 0 Then
'        btn_save.SetFocus
'    Else
'        MsgBox ("Ketahanan barang tidak valid")
'    End If
'End If
'End Sub

Private Sub txt_nama_supplier_LostFocus()
    If Not Me.ActiveControl Is Nothing Then
        If Not Me.ActiveControl.Name = "list_supplier" Then
            list_supplier.Visible = False
        End If
    End If
End Sub

Private Sub txt_nama_supplier_KeyDown(key As Integer, Shift As Integer)
    
    If key = 40 Then
        list_supplier.Visible = True
        list_supplier.SetFocus
        Exit Sub
    ElseIf key = 13 And list_supplier.Visible = True Then
        'txt_kode_supplier = list_supplier.ListItems(0).Text
        'txt_nama_supplier = list_supplier.ListItems(0).SubItems(1)
        list_supplier.SetFocus
    ElseIf key = 13 And list_supplier.Visible = False And txt_kode_supplier.Text <> "" Then
        btn_save.SetFocus
    Else
        txt_kode_supplier = ""
    End If
End Sub

Private Sub list_supplier_LostFocus()
    list_supplier.Visible = False
End Sub

Private Sub list_supplier_dblclick()
    If list_supplier.SelectedItem.index >= 0 Then
        txt_kode_supplier = list_supplier.SelectedItem.Text
        txt_nama_supplier = list_supplier.SelectedItem.SubItems(1)
        'txt_ketahanan.SetFocus
        btn_save.SetFocus
    End If
End Sub

Private Sub list_supplier_keydown(key As Integer, Shift As Integer)
    If key = 13 Then
        list_supplier_dblclick
    End If
End Sub

Private Sub reload_Supplier()
    'list_supplier.Visible = True
    list_supplier.ListItems.Clear
    Dim rsSup As ADODB.Recordset
    Set rsSup = con.Execute("select * from tbsuplier where nmsuplier like '%" & txt_nama_supplier & "%'")
    If rsSup.EOF Then
        list_supplier.Visible = False
        Exit Sub
    End If
    
    rsSup.MoveFirst
    Do While Not rsSup.EOF
        list_supplier.ListItems.Add(, , rsSup!kdsuplier).SubItems(1) = rsSup!nmsuplier
        rsSup.MoveNext
    Loop
    
    Set rsSup = Nothing
End Sub

Private Sub reload_Kategori()
    Dim rsKategori As ADODB.Recordset
    Set rsKategori = con.Execute("select * from tbkategori")
    If Not (rsKategori.EOF Or rsKategori.BOF) Then
        rsKategori.MoveFirst
        Do While Not rsKategori.EOF
            cb_kategori.AddItem (rsKategori!kode)
            rsKategori.MoveNext
        Loop
    End If
End Sub

Private Function cek_Kategori() As Boolean
    cek_Kategori = False
    Dim i As Integer
    Do While i < cb_kategori.ListCount
        If Trim(UCase(cb_kategori.Text)) = Trim(UCase(cb_kategori.List(i))) Then
            cek_Kategori = True
        End If
        i = i + 1
    Loop
End Function

Function Barcode39(InString As String) As String
  ' This function returns the input string with:
  ' a start *, the original string, a check digit, and a stop *
'  Dim i As Integer                ' Counter
'  Dim Chk As Integer              ' Check Digit
'  Dim Char1 As String             ' Current character
  Dim temp As String
'  Dim c39CharSet    As String     ' The 3 of 9 43 character set
'  c39CharSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
'  Chk = 0
'  For i = 1 To Len(InString)
'    Char1 = Mid$(InString, i, 1)
'    ' find the position of this character in the valid set of
'    ' 43 characters and subtract 1 (zero based)
'    Chk = Chk + (InStr(c39CharSet, Char1) - 1)
'  Next i
'
'  temp = Mid$(c39CharSet, (Chk Mod 43) + 1, 1)
  'check digit disabled
  temp = ""
  Barcode39 = "*" & InString & temp & "*"
End Function

Private Sub MonoChrome()
    ' // Convert to B&W //
    Dim DeskWnd As Long, DeskDC As Long
    Dim MyDC As Long
    Dim MyDIB As Long, OldDIB As Long
    Dim DIBInf As BITMAPINFO1
    
    Picture4.AutoRedraw = True

    'Create DC based on desktop DC
    DeskWnd = GetDesktopWindow()
    DeskDC = GetDC(DeskWnd)
    MyDC = CreateCompatibleDC(DeskDC)
    ReleaseDC DeskWnd, DeskDC
    'Validate DC
    If (MyDC = 0) Then Exit Sub
    'Set DIB information
    With DIBInf
        With .bmiHeader 'Same size as picture
            .biWidth = Picture4.ScaleX(Picture4.ScaleWidth, Picture4.ScaleMode, vbPixels)
            .biHeight = Picture4.ScaleY(Picture4.ScaleHeight, Picture4.ScaleMode, vbPixels)
            .biBitCount = 1
            .biPlanes = 1
            .biClrUsed = 2
            .biClrImportant = 2
            .biSize = Len(DIBInf.bmiHeader)
        End With
        ' Palette is Black ...
        With .bmiColors(0)
            .rgbRed = &H0
            .rgbGreen = &H0
            .rgbBlue = &H0
        End With
        ' ... and white
        With .bmiColors(1)
            .rgbRed = &HFF
            .rgbGreen = &HFF
            .rgbBlue = &HFF
        End With
    End With
    ' Create the DIBSection
    MyDIB = CreateDIBSection1(MyDC, DIBInf, 0, ByVal 0&, 0, 0)
    If (MyDIB) Then ' Validate and select DIB
        OldDIB = SelectObject(MyDC, MyDIB)
           BitBlt MyDC, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, Picture4.hdc, 0, 0, vbSrcCopy
        ' Draw the monochome image back to the picture box
        BitBlt Picture4.hdc, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, MyDC, 0, 0, vbSrcCopy
        ' Clean up DIB
        SelectObject MyDC, OldDIB
        DeleteObject MyDIB
    End If
    ' Clean up DC
    DeleteDC MyDC
    ' Redraw
    Picture4.Refresh
End Sub

