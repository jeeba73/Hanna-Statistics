1) On the project properties uncheck the "REMOVE INFORMATION ABOUT UNUSED ACTIVEX CONTROLS".

This is because the scrollbars are added by the scrollAdd control, so the compiler think you never added any scrollbars at all and remove it before executing it.

2) On ucScrollAdd.ctl you need to do 2 aditional changes:

Line 374 and 381 should use the name of the project like:

'Vertical
Set UCScrollV = TargetForm.Controls.Add("Project1.ucScrollBar", UCScrollVName$)
'Horizontal
Set UCScrollH = TargetForm.Controls.Add("Project1.ucScrollBar", UCScrollHName$)