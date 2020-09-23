VERSION 5.00
Begin VB.UserControl NetButton 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   ScaleHeight     =   390
   ScaleWidth      =   1815
   Begin VB.PictureBox picIndi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   1785
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.Label lblIndi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ".Net Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   45
         Width           =   975
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   50
         Top             =   50
         Width           =   240
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00D1ADAD&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
End
Attribute VB_Name = "NetButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Event Declarations:
Event Click() 'MappingInfo=lblIndi,lblIndi,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lblIndi,lblIndi,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Sub UserControl_Resize()
   With picIndi
        .Height = UserControl.ScaleHeight
        .Top = UserControl.ScaleTop
        .Left = UserControl.ScaleLeft
        .Width = UserControl.ScaleWidth
    End With
        With lblIndi
        .Height = UserControl.ScaleHeight
        .Top = (UserControl.ScaleTop / 2) + 45
        .Left = (UserControl.ScaleLeft / 2) + 480
        .Width = UserControl.ScaleWidth
    End With
    Shape1.Height = UserControl.ScaleHeight
    Image1.Height = UserControl.ScaleHeight
    Image1.Width = UserControl.ScaleWidth
End Sub


Private Sub lblIndi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picIndi.BackColor = &HD1ADAD
End Sub

Private Sub lblIndi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picIndi.BackColor = vbWhite
End Sub

Private Sub picIndi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picIndi.BackColor = &HD1ADAD
End Sub

Private Sub picIndi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picIndi.BackColor = vbWhite
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picIndi,picIndi,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = picIndi.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picIndi.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblIndi,lblIndi,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblIndi.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblIndi.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblIndi,lblIndi,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblIndi.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblIndi.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

Private Sub lblIndi_Click()
    RaiseEvent Click
End Sub

Private Sub lblIndi_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblIndi,lblIndi,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblIndi.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblIndi.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Image1.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    picIndi.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    lblIndi.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblIndi.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblIndi.Caption = PropBag.ReadProperty("Caption", ".Net Button")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", picIndi.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", lblIndi.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblIndi.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", lblIndi.Caption, ".Net Button")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

