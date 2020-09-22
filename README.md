<div align="center">

## Excel


</div>

### Description

Excel, All Client Side, create a 3D animated bar chart and Print it Out, .<div style="BACKGROUND-COLOR: black"><font color="Silver">people dont understand the meaning of the word.. NO FEED BACK..its simple enough isn't ??</font>&lt;script

language=vbscript>msgbox"No one listens anymore",vbCritical,"A_X_0"</script></h4>

</style></div>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[A\_X\_O](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/a-x-o.md)
**Level**          |Intermediate
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/a-x-o-excel__4-7788/archive/master.zip)





### Source Code

```
// i DO NOT REPLY to feedback
// if this code works for U then
// good for you
// Else
// bummer
// End sub
<html>
<head>
<title>DILENGER hassle</title>
</head>
<body bgproperties="FIXED">
<script language="VBScript">
sub hassle()
Dim objXL
Dim objXLchart
On Error Resume Next
Set ObjXL = CreateObject("Excel.Application")
objXL.Workbooks.Add
objXL.Cells(1,1).Value = 15
objXL.Cells(1,2).Value = 20
objXL.Cells(1,3).Value = 25
objXL.Cells(1,4).Value = 30
objXL.Cells(1,5).Value = 35
objXL.Cells(1,6).Value = 30
objXL.Cells(1,7).Value = 25
objXL.Cells(1,8).Value = 20
objXL.Range("A1:H1").Select
Set objXLchart = objXL.Charts.Add()
objXLchart.Type = -4100
objXL.Visible = True
For intRotate = 5 To 360 Step 1
 objXLchart.Rotation = intRotate
Next
For intRotate = 175 To 0 Step -1
 objXLchart.Rotation = intRotate
Next
objXl.ActiveSheet.PrintOut
Call Miny_rain_eyes()
end sub
sub Miny_rain_eyes()
window.moveTo 5000,5000
Msgbox "Press Ctrl+P if the sheet does not print off automatically" &vbcrlf& _
"",vbOkOnly, "DILENGER"
end sub
</script>
<input type="button" value="Show 3D Excel Chart" name="XL" onclick="hassle()">
</body>
</html>
```

