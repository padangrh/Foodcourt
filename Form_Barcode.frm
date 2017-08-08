VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Barcode 
   Caption         =   "Cetak Barcode"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   17085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   14160
      Picture         =   "Form_Barcode.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   3795
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3480
      Left            =   13200
      ScaleHeight     =   3420
      ScaleWidth      =   6075
      TabIndex        =   8
      Top             =   2160
      Width           =   6135
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   600
      Picture         =   "Form_Barcode.frx":1AA8
      ScaleHeight     =   1155
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   3480
      Left            =   13200
      Picture         =   "Form_Barcode.frx":3DAA
      ScaleHeight     =   3481.901
      ScaleMode       =   0  'User
      ScaleWidth      =   6075
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   13320
      ScaleHeight     =   1260
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   5880
      Width           =   5895
   End
   Begin VB.TextBox txt_kode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txt_nama 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txt_harga 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   8520
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txt_jumlah 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   614
      Left            =   11400
      TabIndex        =   4
      Top             =   840
      Width           =   1605
   End
   Begin MSComctlLib.ListView list_nama 
      Height          =   2295
      Left            =   3360
      TabIndex        =   2
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode"
         Object.Width           =   2976
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama"
         Object.Width           =   7440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Harga"
         Object.Width           =   2976
      EndProperty
   End
   Begin MSComctlLib.ListView lv_jual 
      Height          =   4935
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Barang"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Kode Supplier"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Jumlah"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "Form_Barcode"
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


Dim rsBarang As ADODB.Recordset
Dim txt_nama_toggle As Boolean
Private CodeClair$, CodeBarre$


Private Sub Form_Load()
'    lbl_user = username
'    txt_total = 0
    kosongkan
    txt_nama_toggle = False
    Set rsBarang = con.Execute("select * from tbbarang")
        
    reload_List

End Sub

Private Sub Form_KeyDown(key As Integer, Shift As Integer)
    If key = 112 Then
        If lv_jual.ListItems.count > 0 Then
'            Form_Print.Show
'            Form_Print.Init lbl_faktur, txt_total, True
'            Me.Enabled = False
            cetakDiKertas
        Else
            MsgBox "Faktur masih kosong"
        End If
    End If
    
    If key = 46 Then
        If Shift = 1 Then
'            txt_total = "0"
            lv_jual.ListItems.Clear
        Else
'            txt_total = Format(priceToNum(txt_total) - priceToNum(lv_jual.SelectedItem.SubItems(4)), "###,###,##0")
            lv_jual.ListItems.Remove (lv_jual.SelectedItem.index)
        End If
    End If
    If key = 115 Then
        If MsgBox("Tutup form transaksi?", vbYesNo) = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub kosongkan()
    txt_kode.Text = ""
    txt_nama.Text = ""
    txt_harga.Text = ""
    txt_jumlah.Text = 1
    list_nama.Visible = False
End Sub

Private Sub list_nama_lostfocus()
    list_nama.Visible = False
End Sub

Private Sub list_nama_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        list_nama_DblClick
    End If
End Sub

Private Sub txt_nama_Change()
    
    If txt_nama.Text <> "" And txt_nama_toggle = False Then
        list_nama.Visible = True
        reload_List
    Else
        list_nama.Visible = False
        txt_nama_toggle = False
    End If
End Sub

Private Sub txt_nama_LostFocus()
    If Not Me.ActiveControl Is Nothing Then
        If Not Me.ActiveControl.Name = "list_nama" Then
            list_nama.Visible = False
        End If
    End If
End Sub

Private Sub list_nama_DblClick()
    If getItemByID(list_nama.SelectedItem.Text) Then
        txt_kode.Text = rsBarang!kode
        txt_nama.Text = rsBarang!nama
        txt_harga.Text = rsBarang!kdsuplier
        list_nama.Visible = False
        txt_jumlah.SetFocus
        txt_jumlah.SelLength = Len(txt_jumlah.Text)
    End If
End Sub

Private Sub txt_Jumlah_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        If Len(txt_jumlah) > 4 Then
            txt_jumlah = ""
            Exit Sub
        End If
    
        If txt_harga = "" Then
            MsgBox "Barang tidak valid"
            Exit Sub
        End If
        
        If Val(txt_jumlah.Text) < 1 Then
            MsgBox "Jumlah tidak valid"
            Exit Sub
        End If
        
        Dim found As Boolean
        Dim i As Integer
        found = False
        i = 1
        
        Do While i <= lv_jual.ListItems.count
            If lv_jual.ListItems(i).Text = rsBarang!kode Then
                found = True
                lv_jual.ListItems(i).SubItems(3) = Val(lv_jual.ListItems(i).SubItems(3)) + Val(txt_jumlah.Text)
                Exit Do
            End If
            i = i + 1
        Loop
        
        If found = False Then
            Dim item As ListItem
            Set item = lv_jual.ListItems.Add(, , rsBarang!kode)
            item.SubItems(1) = rsBarang!nama
            item.SubItems(2) = rsBarang!kdsuplier
            item.SubItems(3) = txt_jumlah.Text
        End If
        
        
        kosongkan
        reload_List
        txt_kode.SetFocus
    End If
End Sub

Private Sub txt_kode_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        txt_nama_toggle = True
        Dim kode As String
        kode = Trim(txt_kode.Text)
        If getItemByID(kode) Then
            txt_nama.Text = rsBarang!nama
            txt_harga.Text = Format(rsBarang!harga_jual, "###,###,###")
            txt_jumlah.SetFocus
            txt_jumlah.SelLength = Len(txt_jumlah.Text)
        Else
            MsgBox ("Kode ini tidak terdaftar")
        End If
    ElseIf Len(txt_nama) > 0 Then
        txt_nama = ""
        txt_harga = ""
    End If
End Sub

Private Function getItemByID(kode As String) As Boolean
    rsBarang.MoveFirst
    Do While Not rsBarang.EOF
        If rsBarang!kode = kode Then
            getItemByID = True
            Exit Function
        End If
        rsBarang.MoveNext
    Loop
    getItemByID = False
End Function

Private Sub txt_nama_KeyDown(key As Integer, Shift As Integer)
    If key = 40 Then
        list_nama.Visible = True
        list_nama.SetFocus
        'Exit Sub
    ElseIf key = 13 And list_nama.Visible = True Then
        list_nama.SetFocus
    End If

End Sub

Public Sub reload_List()
    list_nama.ListItems.Clear
    Dim rsFilter As ADODB.Recordset
    Set rsFilter = con.Execute("select * from tbbarang where nama like '%" & txt_nama.Text & "%'")
    
    If rsFilter.EOF Then
        list_nama.Visible = False
        Exit Sub
    End If
    
    rsFilter.MoveFirst
    Do While Not rsFilter.EOF
        Dim mitem As ListItem
        Set mitem = list_nama.ListItems.Add(, , rsFilter!kode)
        mitem.SubItems(1) = rsFilter!nama
        mitem.SubItems(2) = "Rp. " + Format(rsFilter!harga_jual, "###,###,###")
        rsFilter.MoveNext
    Loop
    
    Set rsFilter = Nothing
End Sub

Sub printBarcode(inKode As String, inNama As String, inSup As String)
    'AddFont (App.Path & "\fre3of9x.ttf")
    AddFont (App.Path & "\code128.ttf")
    
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
    
    Picture4.Cls
    Picture4.AutoRedraw = True
    Picture4.AutoSize = True
    Picture4.FontBold = False
    Picture4.FontSize = 72
'    Picture4.Font = "Free 3 of 9 Extended"
    Picture4.Font = "Code 128"
    Picture4.CurrentY = 10
'    Picture4.CurrentX = (Picture4.ScaleWidth - Picture4.TextWidth(Barcode39(inKode))) / 2
'    Picture4.Print Barcode39(inKode)
    Picture4.CurrentX = (Picture4.ScaleWidth - Picture4.TextWidth(Barcode128$(inKode))) / 2
    Picture4.Print Barcode128$(inKode)

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
    Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(inSup)) * 0.88
    Picture1.Print inSup
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
    Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(inNama)) / 2
    'Picture1.Print Tab(((38 - Len(txt_nama.Text)) / 2) + 1); txt_nama.Text
    Picture1.Print inNama
    'Picture1.Print Tab(2); Barcode39(txt_kode.Text)
    
'    Picture1.Print Tab(((24 - Len(txt_kode.Text)) / 2) + 1); txt_kode.Text
    Picture1.CurrentY = 2500
    Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(inKode)) / 2
    Picture1.Print inKode
    Picture1.Print Tab(1); "                                          "
    Picture1.Print Tab(1); "                                          "
    Picture1.FontSize = 14
'    Picture1.Print Tab(37 - Len(txt_kode_supplier.Text)); txt_kode_supplier.Text
    Picture1.FontBold = False

    Picture1.PaintPicture Picture3.Image, 10, 10


'    Printer.EndDoc
'    RemoveFont (App.Path & "\fre3of9x.ttf")
    RemoveFont (App.Path & "\code128.ttf")
    DoEvents
End Sub

Sub cetakDiKertas()

    Dim xMargin, yMargin As Integer
    Dim resize As Variant
    xMargin = 300
    yMargin = 100
    resize = 0.66
    Dim x, y As Integer
    Dim i As Integer
    Dim posx, posy As Variant
    Dim flagEnd As Boolean
    Dim litem As ListItem
    i = 1
    flagEnd = True
    Dim tempKd As String
    tempKd = ""
    
    
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORLandscape
    Do While flagEnd
        For x = 0 To 3
            posx = x * resize * Picture1.Width + xMargin
            For y = 0 To 4
                If lv_jual.ListItems.count > 0 Then
                    If tempKd <> lv_jual.ListItems(i).Text Then
                        Call printBarcode(lv_jual.ListItems(i).Text, lv_jual.ListItems(i).SubItems(1), lv_jual.ListItems(i).SubItems(2))
                        tempKd = lv_jual.ListItems(i).Text
                    End If

                    posy = y * resize * Picture1.Height + yMargin
                    Printer.PaintPicture Picture1.Image, posx, posy, Picture1.Width * resize, Picture1.Height * resize
                    Printer.PaintPicture Picture4.Image, posx + (Picture1.Width * resize - Picture4.Width * resize * 0.8) / 2, posy + 990, Picture4.Width * resize * 0.8, Picture4.Height * resize * 0.8
'                    Printer.PaintPicture Picture5.Image, posx + (Picture1.Width * resize - Picture5.Width * resize * 0.8) / 2, posy + 1090, Picture5.Width * resize * 0.8, Picture5.Height * resize * 0.8
                    lv_jual.ListItems(i).SubItems(3) = lv_jual.ListItems(i).SubItems(3) - 1
                    If lv_jual.ListItems(i).SubItems(3) = 0 Then
                        lv_jual.ListItems.Remove (i)
                    End If
                    flagEnd = True
                Else
                    flagEnd = False
                    Exit For
                End If
            Next
            If flagEnd = False Then Exit For
        Next
        If flagEnd = True Then
            Printer.NewPage
        Else
            Printer.EndDoc
        End If
    Loop
    
End Sub

Function Barcode39(inString As String) As String
  Dim temp As String
  temp = ""
  Barcode39 = "*" & inString & temp & "*"
End Function

Function Barcode128$(chaine$)
  'Cette fonction est régie par la Licence Générale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'Paramètres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affichée avec la police CODE128.TTF, donne le code barre
  '         * une chaine vide si paramètre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, mini%, dummy%, tableB As Boolean
  Barcode128$ = ""
  If Len(chaine$) > 0 Then
  'Vérifier si caractères valides
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(Mid$(chaine$, i%, 1))
      Case 32 To 126, 203
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculer la chaine de code en optimisant l'usage des tables B et C
    'Calculation of the code string with optimized use of tables B and C
    Barcode128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% devient l'index sur la chaine / i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'Voir si intéressant de passer en table C / See if interesting to switch to table C
          'Oui pour 4 chiffres au début ou à la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub testnum
          If mini% < 0 Then 'Choix table C / Choice of table C
            If i% = 1 Then 'Débuter sur table C / Starting with table C
              Barcode128$ = Chr$(210)
            Else 'Commuter sur table C / Switch to table C
              Barcode128$ = Barcode128$ & Chr$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then Barcode128$ = Chr$(209) 'Débuter sur table B / Starting with table B
          End If
        End If
        If Not tableB Then
          'On est sur la table C, essayer de traiter 2 chiffres / We are on table C, try to process 2 digits
          mini% = 2
          GoSub testnum
          If mini% < 0 Then 'OK pour 2 chiffres, les traiter / OK for 2 digits, process it
            dummy% = Val(Mid$(chaine$, i%, 2))
            dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 105)
            Barcode128$ = Barcode128$ & Chr$(dummy%)
            i% = i% + 2
          Else 'On n'a pas 2 chiffres, repasser en table B / We haven't 2 digits, switch to table B
            Barcode128$ = Barcode128$ & Chr$(205)
            tableB = True
          End If
        End If
        If tableB Then
          'Traiter 1 caractère en table B / Process 1 digit with table B
          Barcode128$ = Barcode128$ & Mid$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calcul de la clé de contrôle / Calculation of the checksum
      For i% = 1 To Len(Barcode128$)
        dummy% = Asc(Mid$(Barcode128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then checksum& = dummy%
        checksum& = (checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calcul du code ASCII de la clé / Calculation of the checksum ASCII code
      checksum& = IIf(checksum& < 95, checksum& + 32, checksum& + 105)
      'Ajout de la clé et du STOP / Add the checksum and the STOP
      Barcode128$ = Barcode128$ & Chr$(checksum&) & Chr$(211)
    End If
  End If
  Exit Function
testnum:
  'si les mini% caractères à partir de i% sont numériques, alors mini%=0
  'if the mini% characters from i% are numeric, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If Asc(Mid$(chaine$, i% + mini%, 1)) < 48 Or Asc(Mid$(chaine$, i% + mini%, 1)) > 57 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
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

