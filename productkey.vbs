Set wshShell = CreateObject("WScript.Shell")
Set writeStream = CreateObject("ADODB.Stream")

outputLicenceKey = ConvertToKey(wshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))

' write file
writeStream.Open
writeStream.Type     = 2 'text
writeStream.Position = 0
writeStream.Charset  = "utf-8"
writeStream.WriteText outputLicenceKey
writeStream.SaveToFile "key.txt", 2
writeStream.Close

MsgBox outputLicenceKey

Function ConvertToKey(Key)
    Const KeyOffset = 52
    i = 28
    Chars = "BCDFGHJKMPQRTVWXY2346789"
	
    Do
      Cur = 0
      x = 14
    
	Do
      Cur = Cur * 256
      Cur = Key(x + KeyOffset) + Cur
      Key(x + KeyOffset) = (Cur \ 24) And 255
      Cur = Cur Mod 24
      x = x -1
    
	Loop While x >= 0
      i = i -1
      KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
    
	If (((29 - i) Mod 6) = 0) And (i <> -1) Then
      i = i -1
      KeyOutput = "-" & KeyOutput
    End If
    
	Loop While i >= 0
      ConvertToKey = KeyOutput
End Function