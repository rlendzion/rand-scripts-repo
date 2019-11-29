''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Create a Dictionary using for loop and randomly select a Key and an Item associated with it
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Randomize
Dim myDict
Set myDict = CreateObject("Scripting.Dictionary")
for x = 0 to 5
myDict.Add "key_"&x , Rnd
Next
items = myDict.Items
'checkpoint
Dim i
i = cInt(Rnd*6)
MsgBox "First item {" & myDict.Keys()(i) & " = " & myDict.Items()(i) & "} out of " & UBound(items)+1 & " elements."
