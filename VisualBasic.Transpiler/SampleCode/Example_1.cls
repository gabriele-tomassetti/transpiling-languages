VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Example_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Nope(1 To 20) As String

Public Type Data
	Num As Integer
	Text(1 To 10) As String
End Type

Public Simple As Data

Public Sub Main_Sub()
    Dim message As String
	Dim result As Integer
	
	Simple.Num = 10
	
	message = Simple.Num & " "
    
	For Index = 0 To UBound(Simple.Text)
        Simple.Text(Index) = " " & Index
		message = message & Simple.Text(Index)
    Next
   
    Debug.Print message   
	
	Call SaveData(message)
	
    result = Multiply(Simple.Num)
	
	Debug.Print result

End Sub
