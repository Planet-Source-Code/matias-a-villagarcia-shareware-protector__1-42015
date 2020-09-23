Attribute VB_Name = "Serial_Generator"
'IMPORTANT NOTE: You can use this code in your
'programs freely but put me in the credits please, and if
'you can let me know.
'Thanks, Matías Ariel Villagarcía.

Public SerialNumber(16) As String 'The 16 characters of the serial n°
Public CompleteSerial As String 'The complete serial number
'////////////////////////////////////////////////////
'I don't know the author's name of this code
Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function GetSerialNumber(strDrive As String) As Long
'This function returns the serial number of the Hard-Disk
Dim SerialNum As Long
Dim Res As Long
Dim Temp1 As String
Dim Temp2 As String
Temp1 = String$(255, Chr$(0))
Temp2 = String$(255, Chr$(0))
Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
GetSerialNumber = SerialNum
End Function
'////////////////////////////////////////////////////

Public Function GenerateSerialHD(Name As String, Drive As String)
Dim HDSerial As Long, TotalName As Integer
Dim Numero As Long, PreSerial As String
'Numero (Number in english) and PreSerial are strings
'because they contain both numbers and letters.
HDSerial = GetSerialNumber(Drive) 'Get Hard-Disk Serial Number
For i = 1 To Len(Name)
    TotalName = TotalName + Asc(Mid(Name, i, 1))
Next i
Numero = HDSerial + TotalName
Numero = Numero Xor HDSerial
Rnd -1
Randomize Numero
'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)
'Int((upper_limit - lower_limit + 1) * Rnd + lower_limit)
For j = 1 To 4
    'Get random numbers
    PreSerial = Int((57 - 48 + 1) * Rnd + 48)
    SerialNumber(j) = Chr(PreSerial)
Next j
For i = 5 To 8
    'Get random Letters
    PreSerial = Int((90 - 65 + 1) * Rnd + 65)
    SerialNumber(i) = Chr(PreSerial)
Next i
For c = 9 To 12
    'Get random Capital Letters
    PreSerial = Int((122 - 97 + 1) * Rnd + 97)
    SerialNumber(c) = Chr(PreSerial)
Next c
For x = 13 To 16
    'Get random numbers
    PreSerial = Int((57 - 48 + 1) * Rnd + 48)
    SerialNumber(x) = Chr(PreSerial)
Next x
CompleteSerial = GetCompleteSerial
GenerateSerialHD = CompleteSerial
End Function

Public Function Check(Name As String, Serial As String, Drive As String) As Boolean
Dim GeneratedSerial As String
GeneratedSerial = GenerateSerialHD(Name, Drive)
If Serial = GeneratedSerial Then
    Check = True
Else
    Check = False
End If
End Function

Public Function GetCompleteSerial()
Dim Serial As String
For i = 1 To UBound(SerialNumber)
    If i = 5 Or i = 9 Or i = 13 Then Serial = Serial & "-"
    Serial = Serial & SerialNumber(i)
Next i
GetCompleteSerial = Serial
End Function
