'datatypes and variables
'one datatype called variant-acts as a string/numeric based on context
'variable declaration - 3 ways-Dim statament, public statement, private statement
Option Explicit 'before declaring variable-allows us to avoid misspell issues
Dim a
a=25
Dim x,y,z
x="sugirtha"
Dim c(2) ' like array, indexing starts from zero
Dim num, str
num=42
str="Mercury"
Dim total
total=100
totl=total+50
msgbox totl 'this will display 150 but totl will be a new variable

'arithmetic: + - * / \: backslash division will give int division, Mod
Dim a,b,sum
'addition
a=10
b=4
sum=a+b
msgbox "addition:" &a&"+"&b&"="&sum
'relational:  = < >  <= >=  <>
'logical AND OR NOT
'control statements: conditional- if else, select case  and loops

Dim num
If True Then
	
End If

'select case
Select
	Case=1
		MSGBOX " "
	case=2	
		msgbox ""
Case Else
	msgbox " "
End select 

'while wend statement

Dim i
i=0
While i<5
	msgbox "i is"&i
	i=i+1
Wend

Dim i
i=0
Do
msgbox " "
i=i+1
Loop while i<5

Do While  i<5
	msgbox "i is" & i
	i=i+1
Loop
'another : do loop until- executes a block of code and repeat the code until condition becomes true

Dim i
i=0
do
msgbox" i"&i
i=i+1
Loop until i=5



