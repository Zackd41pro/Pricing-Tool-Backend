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
    DEV_V1_DEV.ON_Shutdown_V0_01
'code end
End Sub

Private Sub Workbook_Open()
'currently NOT functional as of (8/7/2020) checked by: (Zachary Daugherty)
    'Created By (System)(8/7/2020)
    'Purpose Case & notes:
        'runs other operands on startup
    'Library Refrences required
        'workbook.object
    'Modules Required
        'DEV_VX
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
    'welcome
        DEV_V1_DEV.welcome
        MsgBox ("Need to remove 'unit list' marked range from the DTS refactor code as it has been removed from the data set")
        DEV_V1_DEV.status
        DTS_V1_DEV.status
        matrix_V2.status
        Report_v_Alpha.status
        SP_V1_DEV.status
        String_V1.status
    'run startup checks
        DEV_V1_DEV.On_Startup_V0_01
        DEV_V1_DEV.check_user_in_v0_01 (DEV_V1_DEV.get_username)
    'run DTS Procedure
        MsgBox ("DTS_V1_DEV.run (update_unit_cost): will not run till 'unit list' has been removed from the data set check")
        MsgBox ("workbook_open~DTS_V1_DEV.run (update_unit_cost): failed to run")
        DTS_V1_DEV.run (update_unit_cost)
'code end
 End Sub


