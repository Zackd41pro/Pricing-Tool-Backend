VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
'currently functional as of (8/7/2020) checked by: (Zachary Daughety)
    'Created By (System)(8/7/2020)
    'Purpose Case & notes:
        'runs other operands before close
    'Library Refrences required
        'workbook.object
    'Modules Required
        'DEV_VX
    'Inputs
        'Internal:
            'Na
        'required:
            'NA
        'optional:
            'NA
    'returned outputs
        'NA
'code start
    'dims
        Dim i As Long
    'Log
        Application.DisplayAlerts = False
            On Error Resume Next
                Boots_Report_v_Alpha.Log_Flush (Save_Exit)
            On Error GoTo 0
        Application.DisplayAlerts = True
        ActiveWorkbook.Save
'code end
End Sub

Private Sub Workbook_Open()
    'Created By (System)(12/15/2020)
    'Purpose Case & notes:
        'runs other operands on startup
    'Library Refrences required
        'workbook.object
    'Modules Required
    'Inputs
        'Internal:
            'na
        'required:
            'na
        'optional:
            'na
    'returned outputs
        'na
'code start
    'define variables
        'na
    'run root
        root.On_startup
    'run DTS Procedure
        Call DTS_V2A.run_V0(update_unit_cost)
    'alpha test of DTH
        Call DTH_VA.run_V0(DTH_update_unit_cost, "")
'code end
 End Sub
