VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuat Tanggal Berdasarkan Tipe Interval Waktu"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim TglAwal As String   'Deklarasi variabel
Dim TipeInterval As String
Dim JlhInterval As String
Dim Msg
  On Error GoTo PesanError
  TglAwal = InputBox("Masukkan tanggal awal:", "Tanggal Awal", "22/01/1973")
            'contoh ini, defaultnya 22 Jan 1973
  If StrPtr(TglAwal) = 0 Then Exit Sub
  If Not IsDate(TglAwal) Then
     MsgBox "Tanggal salah!", vbCritical, "Tanggal Tidak Valid"
     Exit Sub
  End If
  
  TipeInterval = InputBox("Masukkan tipe interval " & vbCrLf & "(Pilih salah satu:" & vbCrLf & "d   Jika ingin ditambahkan dengan hari" & vbCrLf & "m   Jika ingin ditambahkan dengan bulan" & vbCrLf & "yyyy Jika ingin ditambahkan dengan tahun)", "Tipe Interval", "m")
      'contoh ini, defaultnya "m" atau bulan
  If StrPtr(TipeInterval) = 0 Then Exit Sub
  If Not (TipeInterval = "d" Or TipeInterval = "m" Or TipeInterval = "yyyy") Then
     MsgBox "Harus d atau m atau yyyy!", vbCritical, "Tipe Salah"
     Exit Sub
  End If
  
  JlhInterval = InputBox("Masukkan jumlah interval yang " & "akan ditambahkan ke Tanggal Awal:", "Jumlah Interval", "100")
             'contoh ini, defaultnya 100
  If Not IsNumeric(JlhInterval) Then
     MsgBox "Harus numerik/angka!", vbCritical, "Tidak Valid"
     Exit Sub
  End If
  Msg = "Tanggal Baru: " & DateAdd(TipeInterval, _
        CInt(JlhInterval), CDate(TglAwal))
  MsgBox Msg, vbInformation, "Tanggal Baru"
  Exit Sub
PesanError:
  MsgBox Err.Number & " - " & Err.Description
End Sub


