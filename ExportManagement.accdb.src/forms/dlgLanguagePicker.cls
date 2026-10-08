Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "App.Products"

Private Sub Form_Open(Cancel As Integer)
    Me.fraLang.DefaultValue = 1
    Me.fraLang.value = 1

    Me.optVN.OptionValue = 1
    Me.optEN.OptionValue = 2
    Me.optBI.OptionValue = 3
End Sub

Private Sub btnOK_Click()
    Me.Visible = False
End Sub

Private Sub btnCancel_Click()
    DoCmd.Close acForm, Me.name
End Sub
