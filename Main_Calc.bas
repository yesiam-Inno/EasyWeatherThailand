Attribute VB_Name = "Main_Calc"
Option Explicit
''''''''''''''''''''''''''Declare for Read/Write Text File'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Global Const CONFIG_FILENAME As String = "config.ini"  'ไฟล์ที่ใช้เก็บค่าชั่วคราว

' ###########   Setup Parameter Name  for Temperature  ###########
Public p_Temp_M As Double
Public p_Temp_C As Double
' ###########   Parameter for Humidity  ###########
Public p_Hum_M As Double
Public p_Hum_C As Double
' ###########   Parameter for Wind  ###########
Public p_Wind_M As Double
Public p_Wind_C As Double
' ###########   Parameter for Rain ###########
Public p_Rain_A As Double
Public p_Rain_B As Double
' ###########   Parameter for Pressure  ###########
Public p_Press_A As Double
Public p_Press_B As Double
Public Sub Main()
' ####   Load Parameter Value  for  Temperature  ###########
p_Temp_M = Val(GetProfileStr("SETTING", "P_TEMP_M", App.Path & "\" & CONFIG_FILENAME))
p_Temp_C = Val(GetProfileStr("SETTING", "P_TEMP_C", App.Path & "\" & CONFIG_FILENAME))
' ####   Load Parameter Value  for  Humidity  ###########
p_Hum_M = Val(GetProfileStr("SETTING", "P_HUM_M", App.Path & "\" & CONFIG_FILENAME))
p_Hum_C = Val(GetProfileStr("SETTING", "P_HUM_C", App.Path & "\" & CONFIG_FILENAME))
' ####   Load Parameter Value  for  Wind  ###########
p_Wind_M = Val(GetProfileStr("SETTING", "P_WIND_M", App.Path & "\" & CONFIG_FILENAME))
p_Wind_C = Val(GetProfileStr("SETTING", "P_WIND_C", App.Path & "\" & CONFIG_FILENAME))
' ####   Load Parameter Value  for  Rain ###########
p_Rain_A = Val(GetProfileStr("SETTING", "P_RAIN_A", App.Path & "\" & CONFIG_FILENAME))
p_Rain_B = Val(GetProfileStr("SETTING", "P_RAIN_B", App.Path & "\" & CONFIG_FILENAME))
' ####   Load Parameter Value  for  Pressure  ###########
p_Press_A = Val(GetProfileStr("SETTING", "P_PRESS_A", App.Path & "\" & CONFIG_FILENAME))
p_Press_B = Val(GetProfileStr("SETTING", "P_PRESS_B", App.Path & "\" & CONFIG_FILENAME))
End Sub

Public Function GetProfileStr(ProgStr As String, KeyNameStr As String, FileName As String) As String
        Dim ReturnStr As String
        Dim NumCnt As Long
        On Error GoTo ErrHandle
'        If KeyNameStr <> "AVAITABLE" And KeyNameStr <> "SELTABLE" Then
            ReturnStr = String(255, 0)
            NumCnt = GetPrivateProfileString(ProgStr, KeyNameStr, "", ReturnStr, 255, FileName)
'        Else 'Value of Keyname is very big,should modify returnStr Length
'            ReturnStr = String(5000, 0)
'            NumCnt = GetPrivateProfileString(ProgStr, KeyNameStr, "", ReturnStr, 5000, FileName)
'        End If
        ReturnStr = Left$(ReturnStr, NumCnt)
        GetProfileStr = ReturnStr
        Exit Function
ErrHandle:
        MsgBox "Error at GetProfileStr function : " & Err.Description, vbCritical, "Error"
        Exit Function
End Function
Public Function WriteRegFile(ByVal sRegFileName As String, ByVal sSection As String, _
ByVal sItem As String, ByVal sText As String) As Boolean
    Dim i As Integer
    On Error GoTo sWriteRegFileError
    
    i = WritePrivateProfileString(sSection, sItem, sText, sRegFileName)
    WriteRegFile = True
    
    Exit Function
sWriteRegFileError:
    WriteRegFile = False
End Function
' Function for Linear
Public Function Cal_Linear(ByVal Lin_M As Double, ByVal Lin_C As Double, ByVal Lin_X As Double) As Double
     Cal_Linear = (Lin_M * Lin_X) + Lin_C
End Function
' Function for Logarithm
Public Function Cal_Logarithm(ByVal Log_A As Double, ByVal Log_B As Double, ByVal Log_X As Double) As Double
     Cal_Logarithm = (Log_A * Math.Log(Log_X)) + Log_B
End Function
' Function for Exponential
Public Function Cal_Exponential(ByVal Exp_A As Double, ByVal Exp_B As Double, ByVal Exp_X As Double) As Double
     Cal_Exponential = Exp_A * Math.Exp(Exp_B * Exp_X)
End Function
