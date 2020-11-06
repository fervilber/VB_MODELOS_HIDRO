VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Printer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const SP As String = " "
Private m_ClientPrinter As Printer
Private Declare Function SendMessageAsLong Lib "user32" _
        Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long _
        ) As Long

Private Declare Function SendMessageAsString Lib "user32" _
                Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As String _
        ) As Long

Private Const EM_GETLINE      As Long = &HC4
Private Const EM_GETLINECOUNT As Long = &HBA
'//ENUMS (Procedure PutText)
Public Enum ptr_PutText
    ptr_LeftToPoint
    ptr_CenterPoint
    ptr_RightToPoint
    ptr_BootomOfPoint
    ptr_TopOfPoint
End Enum

Public Property Set ClientPrinter(v As Object)
    If TypeName(v) = "Printer" Then
       Set m_ClientPrinter = v
    End If
End Property

Private Sub Class_Terminate()
 
     Set m_ClientPrinter = Nothing
End Sub


Public Sub PutText( _
      ByVal x As Single, _
      ByVal y As Single, _
      s As String, _
      Optional AlignmentPointX As ptr_PutText = ptr_RightToPoint, _
      Optional AlignmentPointY As ptr_PutText = ptr_BootomOfPoint _
    )
    With m_ClientPrinter
        Select Case AlignmentPointX
            Case ptr_RightToPoint '//Default
            Case ptr_LeftToPoint: x = x - .TextWidth(s)
            Case ptr_CenterPoint: x = x - .TextWidth(s) / 2
        End Select
        Select Case AlignmentPointY
            Case ptr_BootomOfPoint '//Default
            Case ptr_TopOfPoint: y = y - .TextHeight(s)
            Case ptr_CenterPoint: y = y - .TextHeight(s) / 2
        End Select
       .CurrentX = x
       .CurrentY = y
    End With
    m_ClientPrinter.Print s;
End Sub

Public Sub PutParagraph( _
       ByVal ParagraphLeft As Single, _
       ByVal ParagraphTop As Single, _
       ByVal ParagraphWidth As Single, _
       ByVal ParagraphHeight As Single, _
       ByVal Text As String _
    )
    Dim pos     As Long
    Dim Word    As String
    Dim SumTem  As String
    Dim SumWord As String
    Dim hgt     As Single
    Dim pY      As Single
    Dim s       As String

    hgt = m_ClientPrinter.TextHeight("H")
    s = RTrim$(Text) + SP
    pY = ParagraphTop
    Do
       pos = InStr(s, SP)
       If pos Then
          Word = Left$(s, pos - 1)

          '//Cut word (unfrequent case)
          If m_ClientPrinter.TextWidth(Word) > ParagraphWidth Then
             CutWord Word, ParagraphWidth
          End If

          SumTem = SumTem + Word + SP

          If m_ClientPrinter.TextWidth(SumTem) <= ParagraphWidth Then
             s = Mid$(s, pos + 1)
             SumWord = SumTem
          Else
             PutText ParagraphLeft, pY, SumWord
             pY = pY + hgt
             SumTem = vbNullString
             If (pY + hgt) > (ParagraphTop + ParagraphHeight) Then
                pos = 0 '//Force exit
             End If
          End If
       Else
          PutText ParagraphLeft, pY, SumWord
       End If
    Loop Until pos = 0
End Sub

Private Sub CutWord(ByRef Word As String, dx As Single)
    Dim CutWord As String, i As Integer
    CutWord = vbNullString
    i = 0
    Do
       i = i + 1
       If m_ClientPrinter.TextWidth(CutWord + _
          Mid(Word, i, 1) + "....") < dx Then
          CutWord = CutWord + Mid(Word, i, 1)
       Else
          Exit Do
       End If
    Loop
    Word = CutWord + "..."
End Sub

Public Sub PrintTextBox(txt As TextBox)
   Dim i As Long
   Dim n As Long

   n = GetLineCount(txt)
   For i = 1 To n
       Printer.Print GetLine(txt, i - 1)
   Next
   Printer.EndDoc
End Sub

'// This function fills the buffer with a line of text
'// specified by LineNumber from the text-box control.
'// The first line starts at zero.
Private Function GetLine(txt As TextBox, LineNumber As Long) As String

    '// Scale this to size of text box.
    Const MAX_CHAR_PER_LINE As Long = 80

    Dim ByteLo As Integer
    Dim ByteHi As Integer
    Dim rtn    As Long
    Dim Buffer As String

    ByteLo = MAX_CHAR_PER_LINE And (255) '// [changed 5/15/92]
    ByteHi = Int(MAX_CHAR_PER_LINE / 256) '// [changed 5/15/92]
    Buffer = Chr$(ByteLo) + Chr$(ByteHi) + Space$(MAX_CHAR_PER_LINE - 2)
    rtn = SendMessageAsString(txt.hWnd, EM_GETLINE, LineNumber, Buffer)
    GetLine = Left$(Buffer, rtn)
End Function

'// This function will return the number of lines currently
'// in the text-box control.
Private Function GetLineCount(txt) As Long
    GetLineCount = SendMessageAsLong(txt.hWnd, EM_GETLINECOUNT, 0, 0)
End Function



Public Sub PutImage( _
       ImageLeft As Single, _
       ImageTop As Single, _
       Image As StdPicture, _
       Optional ShowError As Boolean = False _
    )
    On Error GoTo ErrHandler
    m_ClientPrinter.PaintPicture Image, _
                    ImageLeft, ImageTop, _
                    , , , , , , vbSrcAnd
    Exit Sub

ErrHandler:
    If ShowError Then
       MsgBox "Error in Procedire PutImage: " & _
              vbCrLf & Err.Description, vbInformation
    End If
End Sub
Public Sub Rectangle( _
    x1 As Single, y1 As Single, _
    x2 As Single, y2 As Single, _
    Optional Width As Long = 1 _
    )
    m_ClientPrinter.DrawWidth = Width
    m_ClientPrinter.Line (x1, y1)-(x2, y1)
    m_ClientPrinter.Line -(x2, y2)
    m_ClientPrinter.Line -(x1, y2)
    m_ClientPrinter.Line -(x1, y1)
End Sub
Public Sub PintaRectanRelleno( _
    x1 As Single, y1 As Single, _
    x2 As Single, y2 As Single, _
    Optional Width As Long = 1, _
    Optional Color As Long)
    
   If Color = 0 Then Color = rgb(255, 255, 255)
   m_ClientPrinter.DrawWidth = Width
   m_ClientPrinter.Line (x1, y1)-(x2, y2), Color, BF
   
End Sub


Public Sub Imprimir()
   m_ClientPrinter.EndDoc
End Sub

