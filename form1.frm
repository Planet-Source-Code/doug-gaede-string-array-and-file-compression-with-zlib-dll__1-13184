VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Zlib Wrapper Demo"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   5280
      TabIndex        =   5
      Text            =   "9"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compress String (CompressString example)"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   3915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compress Byte Array (CompressByteArray example)"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   3915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compress && Decompress File"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   3915
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   $"form1.frx":0000
      Height          =   1215
      Left            =   4080
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Compression:"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objZlib As clsZlibWrapper

Private Sub Command1_Click()
Dim strFileName As String
Dim strFile As String
Dim strThePath As String

Dim lngResult As Long

strFileName = Dir1.Path
If Right$(strFileName, 1) <> "\" Then strFileName = strFileName + "\"

strThePath = strFileName 'get path
strFile = File1.FileName 'get file name

strFileName = strFileName & strFile 'add file name to path

lngResult = objZlib.CompressFile(strFileName, strFileName & "COMPRESSED", Val(txtLevel))
lngResult = objZlib.DecompressFile(strFileName & "COMPRESSED", strFileName & "DECOMPRESSED")

MsgBox "Done!"

End Sub

Private Sub Command2_Click()

Dim TheBytes() As Byte
Dim cnt As Long
Dim TheChar As Integer
Dim c As Long
Dim res As Long

ReDim TheBytes(102400 - 1) 'allocate precisely 100K

'Fill our bytes from random junk.
'we use 10 bytes of one char, 10 bytes of another,
'and so on.
cnt = 10
For c = 0 To UBound(TheBytes)
    If cnt = 10 Then
      TheChar = Int((255 - 0 + 1) * Rnd + 0)
      cnt = 0
    End If
    TheBytes(c) = TheChar
Next c

MsgBox "Original size: " & CStr(UBound(TheBytes) + 1) & " bytes"

res = objZlib.CompressByteArray(TheBytes(), Val(txtLevel))

MsgBox "Compressed size: " & objZlib.CompressedSize & " bytes"

res = objZlib.DecompressByteArray(TheBytes(), objZlib.OriginalSize)

MsgBox "Decompressed (or should I say original) size: " & CStr(UBound(TheBytes) + 1) & " bytes"

'cleanup
Erase TheBytes

End Sub

Private Sub Command3_Click()
Dim res As Long

Dim OurString As String

OurString = "This is a string. It is just a normal ordinary string."

MsgBox OurString & " [Length: " & Len(OurString) & "]"

res = objZlib.CompressString(OurString, Val(txtLevel))

MsgBox "Compressed string (may not display all characters due to Nulls [Length: " & objZlib.CompressedSize & "]: " & OurString

res = objZlib.DecompressString(OurString, objZlib.OriginalSize)

MsgBox "Here it is, again, decompressed: " & "(" & OurString & ")"

'Note, is res is not 0 then it would mean an error.

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub


Private Sub Form_Load()
Set objZlib = New clsZlibWrapper
Dir1.Path = "c:\"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objZlib = Nothing

End Sub

