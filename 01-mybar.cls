VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mybar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents myevent As VBIDE.CommandBarEvents
Attribute myevent.VB_VarHelpID = -1

Private Sub myevent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim str_code As String
    Const VBE_DIR As String = "\04-github\05-vbe\vbe_code\"
    
    str_code = fso_read_txt(get_my_doc & VBE_DIR & CommandBarControl.Caption)
    insert_code str_code
    CancelDefault = True
End Sub
