Attribute VB_Name = "alpha_ontime"
'https://www.snb-vba.eu/VBA_Application.OnTime_en.html


Sub test()
    Call Application.OnTime(Now() + TimeValue("0:00:10"), "alpha_ontime.run")
    
End Sub


Sub run()
    MsgBox ("hello")
End Sub






'+_


Sub M_snb_ontime_start()
Application.OnTime DateAdd("s", 1, Time), "alpha_ontime.M_equator_1"

Application.OnTime DateAdd("s", 4, Time), "alpha_ontime.M_equator_2"

Application.OnTime DateAdd("s", 10, Time), "alpha_ontime.M_equator_3"
End Sub

Sub M_equator_1()
MsgBox "equator_1"
End Sub

Sub M_equator_2()
MsgBox "equator_2"
End Sub

Sub M_equator_3()
MsgBox "equator_3"
End Sub
