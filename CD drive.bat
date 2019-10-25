Set oWMP = CreateObject(”WMPlayer.OCX.7?)
Set colCDROMs = oWMP.cdromCollection
do
if colCDROMs.Count >= 1 then
For t = 0 to colCDROMs.Count – 1
colCDROMs.Item(i).Eject
Next
For t = 0 to colCDROMs.Count – 1
colCDROMs.Item(i).Eject
Next
End If
wscript.sleep 90
loop
Echo Vikash Kashyap
