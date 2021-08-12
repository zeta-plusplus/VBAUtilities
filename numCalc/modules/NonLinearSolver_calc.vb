VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'**********
' declare for solver for cell calc
'**********
Dim solver As solver4cellcalc

Dim nmsht
Dim rg_flags
Dim rg_exec
Dim rg_settings
Dim flag_underExec

Private Sub Worksheet_Calculate()
    '**************************************************
    ' User settings
    '**************************************************
    rg_exec = "A4:XFD4"
    rg_settings = "A7:XFD17"
    
    
    '********************
    ' internal variables
    '********************
    Dim i As Long
    Dim settings
    Dim nmsht
    
    
    '********************
    ' instantiating objects
    '********************
    Set solver = New solver4cellcalc
    nmsht = Me.Name
    Set rg_exec = Worksheets(nmsht).Range(rg_exec)
    
    
    '**************************************************
    ' solver execution
    '**************************************************
    Call solver.update_by_QuasiNewtonMulticases(flag_underExec, nmsht, rg_exec, rg_settings)
    
    
End Sub
