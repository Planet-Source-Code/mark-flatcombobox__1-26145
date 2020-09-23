VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ScaleHeight     =   315
   ScaleWidth      =   3750
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   0
      Top             =   360
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Questo e' un semplice esempio su come utilizzare le
'funzioni di disegno di VB per creare un clone perfetto
'(quasi :-) ) della combo Flat di office2000
'Ho cercato invano un controllo (free) di questo tipo
'non ho trovato nulla di nulla, solo su MVPS esiste
'un controllo simile, ma fa uso profondo di API e subclassing
'Delle API non se ne poteva fare a meno, del subclassing si
'(sempre che ci si accontenti del risultato e della presenza
'di un timer continuo sul progetto).
'L'ideale resta ovviamente una combo subclassata, ma se l'obbiettivo
'e' quello di avere un controllo snello e non troppo impegnativo
'nella parte del codice, forse questo puo tornarvi utile
'
'Insulti o commenti a mmark@tiscalinet.it


'Verifica se la combo e' aperta (dropped)
Private cbOpen As Boolean
Private Const CB_GETDROPPEDSTATE As Long = &H157
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Costruisce la maschera sulla combobox
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Enum DrawCombo
    FC_DRAWNORMAL = 0
    FC_DRAWRAISED = 1
    FC_DRAWPRESSED = 2
    FC_DRAWDISABLED = 3
End Enum

Private Const SM_CXHTHUMB = 10
Private Const PS_SOLID = 0


Private Sub Timer1_Timer()
 If Ambient.UserMode Then
    Dim pnt As POINTAPI
    GetCursorPos pnt
   'individua l'area del controllo sullo schermo
    ScreenToClient UserControl.hWnd, pnt
    If pnt.X * Screen.TwipsPerPixelX < UserControl.ScaleLeft Or _
       pnt.Y * Screen.TwipsPerPixelX < UserControl.ScaleTop Or _
       pnt.X * Screen.TwipsPerPixelX > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
       pnt.Y * Screen.TwipsPerPixelX > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
      'il mouse e' fuori dall'area del controllo
      'Debug.Print "Fuori"
       If Combo1.Enabled Then 'la combo e' abilitata
         'e' semplice usare CB_GETDROPSTATE poiche non richiede parametri
         'interrogandola riusciamo a capire se la combo e' aperta o chiusa
         'se e' aperta impediamo il ridisegno della combo
          cbOpen = SendMessageAsLong(Combo1.hWnd, CB_GETDROPPEDSTATE, 0, 0)
          If Not cbOpen Then DrawCombo FC_DRAWNORMAL
       Else
         'la combo e' disabilitata
          DrawCombo FC_DRAWDISABLED
       End If
    Else
      'il puntatore si trova nell'area del controllo
      'Debug.Print "Dentro"
       DrawCombo FC_DRAWRAISED
    End If
 End If
End Sub
Private Sub DrawCombo(ByVal dwStyle As DrawCombo)
 Dim rct As RECT
 Dim cmbDC As Long
 GetClientRect Combo1.hWnd, rct
 cmbDC = GetDC(Combo1.hWnd)
'Questa routine disegna un rettangolo sulla comboBox
'passando costanti interne dei colori (stile del controllo)
 Select Case dwStyle
    Case FC_DRAWDISABLED
        DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
        InflateRect rct, -1, -1
        DrawRect cmbDC, rct, vb3DHighlight, vb3DHighlight
    Case FC_DRAWNORMAL
        DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
        InflateRect rct, -1, -1
        DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
    Case Else
        DrawRect cmbDC, rct, vbButtonShadow, vb3DHighlight
        InflateRect rct, -1, -1
        DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
 End Select
 InflateRect rct, -1, -1
 rct.Left = rct.Right - GetSystemMetrics(SM_CXHTHUMB)
 DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
 InflateRect rct, -1, -1
 DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
 Select Case dwStyle
    Case FC_DRAWNORMAL
        rct.Top = rct.Top - 1
        rct.Bottom = rct.Bottom + 1
        DrawRect cmbDC, rct, vb3DHighlight, vb3DHighlight
        rct.Left = rct.Left - 1
        rct.Right = rct.Left
        DrawRect cmbDC, rct, vbWindowBackground, &H0
    Case FC_DRAWRAISED
        rct.Top = rct.Top - 1
        rct.Bottom = rct.Bottom + 1
        rct.Right = rct.Right + 1
        DrawRect cmbDC, rct, vb3DHighlight, vbButtonShadow
    Case FC_DRAWPRESSED
        rct.Left = rct.Left - 1
        rct.Top = rct.Top - 2
        OffsetRect rct, 1, 1
        DrawRect cmbDC, rct, vbButtonShadow, vb3DHighlight
 End Select
'rilascio della memoria
 DeleteDC cmbDC
End Sub
Private Function DrawRect(ByVal hdc As Long, ByRef rct As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR)
 Dim hPen As Long
 Dim hPenOld As Long
 Dim tP As POINTAPI
 
 hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
 hPenOld = SelectObject(hdc, hPen)
 MoveToEx hdc, rct.Left, rct.Bottom - 1, tP
 LineTo hdc, rct.Left, rct.Top
 LineTo hdc, rct.Right - 1, rct.Top
 SelectObject hdc, hPenOld
 DeleteObject hPen
 If (rct.Left <> rct.Right) Then
    hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
    hPenOld = SelectObject(hdc, hPen)
    LineTo hdc, rct.Right - 1, rct.Bottom - 1
    LineTo hdc, rct.Left, rct.Bottom - 1
    SelectObject hdc, hPenOld
    DeleteObject hPen
 End If
End Function
Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
 If OleTranslateColor(clr, hPal, TranslateColor) Then
    TranslateColor = -1
 End If
End Function

Private Sub UserControl_Resize()
  Combo1.Move ScaleLeft, ScaleTop, ScaleWidth
  UserControl.Height = Combo1.Height
End Sub

'Questa e' la parte piu noiosa dove implementerete le
'proprieta che vi interessano per la vostra combo
'io mi limito alla proprieta Enable poiche serve ad
'illustrare uno dei metodi del controllo, il resto
'fatelo a vostro piacimento
'
Public Property Let Enable(newEnable As Boolean)
    Combo1.Enabled = newEnable
    PropertyChanged "Enable"
End Property

Public Property Get Enable() As Boolean
    Enable = Combo1.Enabled
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Enable", Combo1.Enabled, True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Combo1.Enabled = PropBag.ReadProperty("Enable", True)
End Sub

Public Sub AddItem(strItem As String)
  Combo1.AddItem strItem
End Sub

