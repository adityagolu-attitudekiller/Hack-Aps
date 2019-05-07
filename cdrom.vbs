set oWMP = Createobject("WMPlayer.ocx.7")
set colCDROMs = oWMP.cdromcollection
if colCDROMs.count >= 1 then
for i = 0 to colCDROMs.count -1
colCDROMs.Item(i).Eject
Next ' cdrom
end if