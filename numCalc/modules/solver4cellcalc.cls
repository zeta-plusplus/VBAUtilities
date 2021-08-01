VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "solver4cellcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------
' update: 2021/07/24(Sat)
' opened to public on repository: https://github.com/zeta-plusplus/VBAUtilities
' rev.001
'--------------------------------------------------------------------------------


Public Sub update_by_QuasiNewtonMulticases(flag_underExec, nmsht, rg_exec, rg_settings)
    '----------------------------------------
    ' declare variables
    '----------------------------------------
    Dim i As Long
    Dim settings
    Dim rg_settings_current
    Dim settings_current
    
    
    
    '----------------------------------------
    ' instantiate objects
    '----------------------------------------
    Set settings = Worksheets(nmsht).Range(rg_settings)
    
    
    
    '----------------------------------------
    ' process
    '----------------------------------------
    For i = 1 To rg_exec.Columns.Count Step 1
        
        
        
        If (rg_exec(1, i) = "exec") And (flag_underExec <> "UnderExec") Then
            Set settings_current = settings.Columns(i)
            rg_settings_current = settings_current.Address
            
            Call update_by_QuasiNewton(flag_underExec, nmsht, _
                                        rg_settings_current)
        End If
        
    Next i
    
    
End Sub


Public Sub update_by_QuasiNewton(flag_underExec, nmsht, rg_settings)
    
    '----------------------------------------
    ' declare variables
    '----------------------------------------
    Dim infoSettings
    
    Dim rg_switches
    Dim rg_flags
    Dim rg_x_n
    Dim rg_f_n
    Dim rg_dx
    Dim rg_Ulim_delta_x
    Dim rg_delta_x
    Dim rg_x_prev
    Dim rg_f_prev
    Dim rg_delta_f
    Dim rg_Jacobian
    
    Dim switches
    Dim flags
    Dim x_n
    Dim f_n
    Dim dx
    Dim Ulim_delta_x
    Dim delta_x
    Dim x_next
    Dim f_next
    Dim delta_f
    Dim x_n_save
    Dim f_n_save
    Dim df
    Dim Jacobian
    Dim detJacobian
    Dim invJacobian
    Dim matTemp
    
    Dim sizeSys
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim errMsg
    
    errMsg = ""
    infoSettings = Worksheets(nmsht).Range(rg_settings)
    
    
    '----------------------------------------
    ' define range objects
    '----------------------------------------
    i = 1
    Set rg_switches = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_dx = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_Ulim_delta_x = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_flags = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_x_n = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_f_n = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_delta_x = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_x_prev = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_f_prev = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_delta_f = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    Set rg_Jacobian = Worksheets(nmsht).Range(infoSettings(i, 1)): i = i + 1
    
    '----------------------------------------
    ' read cells values
    '----------------------------------------
    switches = rg_switches.Cells
    flags = rg_flags.Cells
    x_n = rg_x_n.Cells
    f_n = rg_f_n.Cells
    dx = rg_dx.Cells
    Ulim_delta_x = rg_Ulim_delta_x.Cells
    delta_x = rg_delta_x.Cells
    delta_f = rg_delta_f.Cells
    
    sizeSys = rg_x_n.Count
    ReDim Jacobian(1 To sizeSys, 1 To sizeSys)
    ReDim df(1 To sizeSys, 1 To 1)
    ReDim delta_f(1 To sizeSys, 1 To 1)
    
    '----------------------------------------
    ' calc
    '----------------------------------------
    flag_underExec = "UnderExec"
    
    flags(1, 1) = "UnderExec"
    flags(2, 1) = "NotEvaluated"
    rg_flags.Cells = flags
    
    x_n_save = x_n
    f_n_save = f_n
    x_next = x_n
    
    If (switches(2, 1) = "Newton") Then
        '***** small purturbation obtain jacobian *****
        'loop about x
        For i = 1 To sizeSys Step 1
            x_n(i, 1) = x_n(i, 1) + dx(i, 1)
            rg_x_n.Cells(i, 1) = x_n(i, 1)
            
            If (switches(1, 1) = "ThisSheet") Then
                Worksheets(nmsht).Calculate
            ElseIf (switches(1, 1) = "EntireBook") Then
                Application.Calculate
            End If
            
            f_n = rg_f_n.Cells
            
            'loop about f
            For j = 1 To sizeSys Step 1
                df(j, 1) = f_n(j, 1) - f_n_save(j, 1)
                Jacobian(j, i) = df(j, 1) / dx(i, 1)
                
            Next j
            
            rg_x_n.Cells = x_n_save
                
            If (switches(1, 1) = "ThisSheet") Then
                Worksheets(nmsht).Calculate
            ElseIf (switches(1, 1) = "EntireBook") Then
                Application.Calculate
            End If
            
        Next i
        errMsg = ""
    Else
        errMsg = "inappropriate method setting"
    End If
    
    k = 1
    
    'loop about f
    For i = 1 To sizeSys Step 1
        'loop about x
        For j = 1 To sizeSys Step 1
            rg_Jacobian.Cells(k, 1) = Jacobian(i, j)
            k = k + 1
        Next j
    Next i
    
    detJacobian = WorksheetFunction.MDeterm(Jacobian)
    
    If (Abs(detJacobian) > 0) Then
        flags(2, 1) = "invertible"
        rg_flags.Cells = flags
        
        '*****
        invJacobian = WorksheetFunction.MInverse(Jacobian)
        
        ReDim matTemp(1 To sizeSys, 1 To sizeSys)
        For i = 1 To sizeSys Step 1
            For j = 1 To sizeSys Step 1
                matTemp(i, j) = (-1#) * invJacobian(i, j)
            Next j
        Next i
        
        delta_x = WorksheetFunction.MMult(matTemp, f_n_save)
        
        For i = 1 To sizeSys Step 1
            If (delta_x(i, 1) > Ulim_delta_x(i, 1)) Then
                delta_x(i, 1) = Ulim_delta_x(i, 1)
            ElseIf (delta_x(i, 1) < -1# * Ulim_delta_x(i, 1)) Then
                delta_x(i, 1) = -1# * Ulim_delta_x(i, 1)
            End If
            
            x_next(i, 1) = x_n(i, 1) + delta_x(i, 1)
            
        Next i
        
        '*****
        rg_x_n.Cells = x_next
        If (switches(1, 1) = "ThisSheet") Then
            Worksheets(nmsht).Calculate
        ElseIf (switches(1, 1) = "EntireBook") Then
            Application.Calculate
        End If
        f_next = rg_f_n.Cells
        
        For i = 1 To sizeSys Step 1
            delta_f(i, 1) = f_next(i, 1) - f_n_save(i, 1)
        Next i
        
        '*****
        rg_delta_x.Cells = delta_x
        rg_x_prev.Cells = x_n_save
        rg_f_prev.Cells = f_n_save
        rg_delta_f.Cells = delta_f
    End If
    
    rg_x_n.Cells = x_next
    If (switches(1, 1) = "ThisSheet") Then
        Worksheets(nmsht).Calculate
    ElseIf (switches(1, 1) = "EntireBook") Then
        Application.Calculate
    End If
    
    flag_underExec = "NotUnderExec"
    
    flags(1, 1) = "NotUnderExec"
    rg_flags.Cells = flags
    
End Sub

