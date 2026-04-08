Attribute VB_Name = "xlAppScript_demos"
'//--------------------------------------------------------------------------------------------------------------------------------
'//     Method 1
'//
'//Run xlAppScript from VBA
'
Sub method_1()

'Declare variable to pass through code
Dim article As String

'Setup workbook if this is your first time running xlAppScript
Call connectWb

'Enter some xlAppScript code & set to Lexer variable
article = "<lib>xbas;rng(A1).value(Testing123).bgcolor(gainsboro).fcolor(cornflowerblue);$"

'Send code to the xlAppScript Lexer
Call xlas(article)

'/All done!
'//_______________________________________________________________________________________________________________________________
End Sub
'//     Method 2
'//
'//Run xlAppScript from a file
'
Sub method_2()

'Declare variable to pass through code
Dim article As String

'Setup workbook if this is your first time running xlAppScript
Call connectWb

'Set file containing xlAppScript to a variable
file = Environ("USERPROFILE") & "\documents\demo.txt"

Open file For Input As #1 'open file
Do Until EOF(1) 'search until the end
Line Input #1, filetext 'set code in file to variable
textstring = textstring & filetext 'concatenate text string
Loop
Close #1

article = textstring 'set article to text string

'Send code to the xlAppScript Lexer
Call xlas(article)

'/All done!
'//_______________________________________________________________________________________________________________________________
End Sub
'//     Method 3
'//
'//Send xlAppScript to another connected environment & run
'/(code can be adjusted for parsing through a file as seen in method 2)
'
Public Sub method_3()

'Declare variable to pass through code
Dim article As String

'Set workbook active, & other lines of xlas code (this workbook was located in \documents)
article = "<lib>xbas;wb(xlasbook.xlsm).active;rng(A1).value(Testing123).bgcolor(gainsboro).fcolor(cornflowerblue);$"

'Send code to workbook for parsing using VBA (this will open the workbook)
X = Application.Run("'xlasbook.xlsm'!xlas", (article))

'/All done!
'//_____________________________________________________________________________________________________________
End Sub
