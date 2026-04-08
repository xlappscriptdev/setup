======================================================================
•▀			    xlAppScript Setup Guide		            
======================================================================
Written by: anz7re (André)

----------------------------------------------------------------------
Latest Revision:

1/31/2023

----------------------------------------------------------------------
Developer(s): 

anz7re (André)

----------------------------------------------------------------------
Contact:

Email: support@xlappscript.org | support@autokit.tech | anz7re@autokit.tech
Social: @anz7re | @xlAppScript | @AutokitTech 
Web: xlappscript.org | autokit.tech/xlappscript
Donate: $donateautokitdevs

(Don't hesitate to reach out if you're having any issues!)

/====================================================================================================================\
xlAppScript is a modifiable, automation scripting tool namely for Microsoft Excel (VBA), Windows OS, & Autokit applications.
/====================================================================================================================/

License Information:

Copyright (C) 2022-present, Autokit Technology.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:


1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, 
THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES 
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) 
HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=====================================================================================================================\
INSTALL GUIDE:
=====================================================================================================================/

1. Import the "xlAppScript_lex.bas" & "xlAppScript_setup.bas", along w/ the xlAppScript librarie(s) of your choice 
into your Excel workbook (we recommend atleast importing the "xbas" library for now)

2. Run the "connectWb" macro

All done!

(If you'd like to use a Run Tool, import the corresponding ".frm" file into your workbook.
***Make sure the ".frx" file is in the same location as the ".frm" file)

=====================================================================================================================
XLAPPSCRIPT GUIDE: Getting Started: Choosing a Run Method					
=====================================================================================================================

Now that you have everything installed it's time to write your first lines!

You can write xlAppScript in any text editor, but in order for it to run you'll need to have it
parsed through VBA. 

There's a few different ways to go about doing this:

First let's open up Excel, create a blank worksheet, & head over to the VBA editor.

From here create a new module (the name doesn't matter for now).

---------------------------------------------------------------------------------------------------------------------
Copy & paste "Example 1" into your VBA editor:
---------------------------------------------------------------------------------------------------------------------
# Example 1 #
---------------------------------------------------------------------------------------------------------------------
Sub xlAppScriptDemo1()

'Declare variable to pass through code
Dim article As String

'Setup workbook if this is your first time running xlAppScript
Call connectWb

'Enter some xlAppScript code & set to Lexer variable
article = "<lib>xbas;rng(A1).value(Testing123).bgcolor(gainsboro).fcolor(cornflowerblue);$"

'Send code to the xlAppScript Lexer
Call xlas(article)

'All done!

End Sub
---------------------------------------------------------------------------------------------------------------------

Using the above example as a template, xlAppScript code can easily be placed within the "article"
variable. The contents of "article" are then sent through for lexing/parsing.

This brings us to the run initializer or "$" symbol which must be included somewhere within the
xlAppScript code to initiate lexing/parsing (best practice is just placing it at the end).

Okay let's save some xlAppScript code to a file for later use... 

Copy: "<lib>xbas;kin(v=value);rng(A1).@v(101);$"

Open an empty text document & paste your xlAppScript code. Save this file to your 
documents folder as "demo" for now (you can use ".txt" or ".xlas" for xlAppScript). 

On to our next example...
---------------------------------------------------------------------------------------------------------------------
Copy & paste "Example 2" into your VBA editor:
---------------------------------------------------------------------------------------------------------------------
# Example 2 #
---------------------------------------------------------------------------------------------------------------------
Sub xlAppScriptDemo2()

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

'All done!
End Sub
---------------------------------------------------------------------------------------------------------------------

Using the above example xlAppScript can be parsed from a designated file by retrieving the
code contents, setting them to our "article" variable, & then sending that through the lexer/parser.

This is probably the better method as it allows you more flexibibilty since you can
premeditatively create your scripts then write another script that calls your VBA
macro outside of Excel (which will run the xlAppScript located in your file).

***Aside from these 2 methods shown above, you can also use Autokit application "FlowStrips" (that bright green bar) or the 
"Control Box" console application(s) to trigger xlAppScript code.

---------------------------------------------------------------------------------------------------------------------

The official xlAppScript guidebook is in the works & coming soon!
