VERSION 5.00
Begin VB.Form PrintingTuitorial 
   Caption         =   "VB and Printing"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "PrintingTuitorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TimesClicked As Integer
Dim AnArray(0 To 29)    'Declare the Array





'Extra Notes:
' When I first looked on the Net to learn and "conquer" printing in VB, I found
' very little other than Harvest R's tutorial on PSC that said anything other than
' "use PrintForm" which is OK if your form looks just like the info you want to print.
' In fact, most articals on the subject leaned towards stateing that VB wasn't very strong
' when it came to the printed page. My first reaction was "Have I chosen the right
' language to learn?" but I now think VB can do anything your imagination
' can concoct.
' If you look closely at the code you will find it is not in sequential order.
' I have found Visual Basic can work it out anyway, although the code "may"
' execute faster if you take the time to help VB along.
' You can use this ability to group all the BOLD, Boxes, Graphics, Loops etc
' statements together so VB has less processing to do (I havn't done so much of that here).
' Writing a printing sub can be time consumming and also has it's fair share
' of trial and error. At times you can add a statement and a previously working
' section of code stops working and it is a matter of moving the statement to a different
' place to get it working again.
' The best method I have found to create your page/s is to draw one up by hand
' as a plan, use a ruler, and take as much guesswork out of it as possible
' the goal being to have a finished and functional end product.
' There is much more in VB to handle any printing project than I have looked at,
' hopefully this is a start and a help. ***Some parts of this code are from Harvest R***
' Have Fun.


Private Sub Command1_Click()

    Dim Line1 As String
    Dim Line2 As String
    Dim Line3 As String
    Dim Line4 As String
    Dim Line5 As String
    Dim Line6 As String
    Dim CenteredText As String, CenteredTextWidth As Single
    Dim Answer As String
    Dim TodaysDate As Variant

    Dim HorizontalMargin As Long, VerticalMargin As Long
    Dim Col(0 To 3), NR         '  4 Columns and Next Row
    Dim Col2(0 To 2), NC       ' 3 Columns and Next Column
    Dim f As Integer
    Dim k As Integer
    Dim k2 As Integer
    Dim M
    Dim Units
    Dim ListWord


Select Case TimesClicked        'This is what I want the Select Case to look at each time a user clicks the command button

        Case Is = 0
            Line1 = "Like any tutorial, this one won't tell the complete "
            Line2 = "story, it also won't tell you every way in which a"
            Line3 = "particular VB function can work.  Only the way I"
            Line4 = "make VB work for me. This tutorial is primarily"
            Line5 = "about printing, a beginner may also learn a few things"
            Line6 = "about Select Case, Labels, Text boxes etc."
            
            Label1.Caption = "Welcome"
            Label2.Caption = Line1 & Chr(10) & Line2 & Chr(10) & Line3 & Chr(10) & Line4 & Chr(10) & Line5 & Chr(10) & Line6
                                    
            Command1.Caption = "Next"
            TimesClicked = TimesClicked + 1
        
        Case Is = 1
            Line1 = "After you have progressed through to the end"
            Line2 = "you can see what has been printed on paper and"
            Line3 = "match each piece with the code that put it there."
            Line4 = "A Sub for printing can be very long and the more "
            Line5 = "complex you make the end product the longer it will be."
            
            
            Label1.Caption = "Coding"
            Label2.Caption = Line1 & Chr(10) & Line2 & Chr(10) & Line3 & Chr(10) & Line4 & Chr(10) & Line5
            
            TimesClicked = TimesClicked + 1
        
        Case Is = 2
            Line1 = "We will look at how to print a simple graphic"
            Line2 = "in the position you want it. Graphics also include"
            Line3 = "lines on the paper, vertical, horizontal, short,"
            Line4 = "long, thick and thin. We will print the picture"
            Line5 = "below in two places and two sizes, a circle,"
            Line6 = "a rectangle and a dot."
            
            Label1.Caption = "Graphics"
            Label2.Caption = Line1 & Chr(10) & Line2 & Chr(10) & Line3 & Chr(10) & Line4 & Chr(10) & Line5 & Chr(10) & Line6
            
            Picture1.Visible = True
            
            TimesClicked = TimesClicked + 1
        
        Case Is = 3
            Line1 = "We will also show how to print text and information"
            Line2 = "from anywhere in your program to anywhere on"
            Line3 = "paper. We will print in rows and then in columns."
            Line4 = "It will be the same information from an array, directly "
            Line5 = "and from a listbox, set out differently and using a loop "
            Line6 = "to accomplish both.  Ready?"
            Command1.Caption = "Print Now"
            Label1.Caption = "Text"
            Label2.Caption = Line1 & Chr(10) & Line2 & Chr(10) & Line3 & Chr(10) & Line4 & Chr(10) & Line5 & Chr(10) & Line6
            
            Text1.Text = "We will put this at the bottom"
            Text2.Text = "This one we'll put on the top"
            Text3.Text = "This one in a box"
            
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = True
             List1.Visible = True
            
            TimesClicked = TimesClicked + 1
        
        
        
        Case Is = 4
        
        
                Answer = MsgBox("Confirm printing on " & Printer.DeviceName, vbYesNo) ' This will check who is the system
                If Answer = vbNo Then GoTo Cancel 'The user pressed cancel                  'default printer, it does not use the Common Dialog
            '****************************************************************************************************************
            '****************************************************************************************************************

            '   Dim BeginPage, EndPage, NumCopies, i
            '    ' Set Cancel to True
            '    CommonDialog1.CancelError = True
            '    On Error GoTo ErrHandler
            '    ' Display the Print dialog box
            '    CommonDialog1.ShowPrinter
            '    ' Get user-selected values from the dialog box
            '    BeginPage = CommonDialog1.FromPage
            '    EndPage = CommonDialog1.ToPage
            '    NumCopies = CommonDialog1.Copies
            '    For i = 1 To NumCopies
            '        ' Put code here to send data to the printer
            '
            '    Next i
            '    Exit Sub
            'ErrHandler:
            '    ' User pressed the Cancel button
            '    Exit Sub
            '****************************************************************************************************************
            'This is a sample out of VB Help on how to use the Common Dialog
            'What we will be looking at is more to do with sending data to the printer.
            '
            '
            '****************************************************************************************************************
            '****************************************************************************************************************
            Text1.Visible = False
            Text2.Visible = False
            Text3.Visible = False
             List1.Visible = False
            
            Line1 = "Well done!"
            Line2 = "The Tutorial will now be printed."
            Line3 = "Check out the code, change it and see"
            Line4 = "what happens when you do. I hope this"
            Line5 = "will be helpful now and for future reference."
            
            
            Printer.ScaleMode = vbMillimeters
            'This time I measure in millimeters
            
            
            'This could also be writen as
            'Printer.ScaleMode = 6
            'Check the ScaleMode Constants in VB Help to choose a measurement that suits you.
            
            '{{{{One of the handiest tools you will use while coding to send data from your
            'application to a printer is the common ruler (or some form of measuring device)
            'you will use it often to check and change your setout on paper.}}}}
            
            
            
            'This page will print on A4 format paper (210 * 297 millimeters)
            'This next statement checks the printable area set in the printer properties
            'and deducts that figure from the page width, then divides the remainder by 2
            
            HorizontalMargin = (210 - Printer.ScaleWidth) / 2
            'Debug.Print HorizontalMargin
            VerticalMargin = (297 - Printer.ScaleHeight) / 2
            'Debug.Print VerticalMargin
            
            'Here 5 mm is added horizontally and
            '5 mm vertically in your non printable area, in addition to the previous test
            'because the previous check doesn't take into account that someone may
            'have set the properties of their printer to a zero border.
            
            HorizontalMargin = 5 + HorizontalMargin
            'Debug.Print HorizontalMargin
            VerticalMargin = 5 + VerticalMargin
            'Debug.Print VerticalMargin
            
            'So if the first test showed the border as 6 mm all round,
            'now it would be 11 mm.
            'But if it was zero, there is only 5 mm. and if it was set at 25 mm then it could cause problems
            'later.  You could get around this with some "If HorizontalMargin < 3 Then HorizontalMargin = 5 + HorizontalMargin"
            'statements, it depends how accurate you need to be or how much work you want to do.
                        
                        'We must first define the font and attributes of the text
            Printer.FontName = "Arial"
            Printer.FontSize = 12
            Printer.FontBold = True          'we want bold   ***** NOTE : If you don't want the rest of
            Printer.FontItalic = False       'no italic           ***** your document to be printed in BOLD / Italic etc
            Printer.FontUnderline = False    'no underline ***** then you must "turn off" these switches
            Printer.FontStrikethru = False   'no strike       ***** when your finished.
            Printer.ForeColor = RGB(0, 0, 0) 'color black
            Printer.FillStyle = 1                    'Transparent

            'initialize the printer
            Printer.Print ""
            
            TodaysDate = Format(Date, "Long Date")

Printer.CurrentX = 25
Printer.Print "Digit's VB Printing tutorial"; Space(55);
Printer.Print TodaysDate

            'We'll now draw a line down both sides of our printable area. We use the paper size
            '(210*297) and we add the margins to the starting point
            
            '       NOTE: the syntax is: Printer.Line (X1,Y1)-(X2,Y2),color,flag
            '                   EG:   Printer.Line {Starting Point} ( X1 is from Left Margin, Y1 is from Top Margin)-( X2 is from Left Margin, Y2 is from Top Margin){Finishing Point}
            '             where flag can be: nothing (draw a line),
            '                                B       (draw a box) or
            '                                BF      (draw a filled box)
            Printer.Line (HorizontalMargin, VerticalMargin)-(HorizontalMargin, 280 - VerticalMargin), RGB(0, 0, 0)   'Left hand side
            Printer.Line (174 + HorizontalMargin, VerticalMargin)-(174 + HorizontalMargin, 280 - VerticalMargin), RGB(0, 0, 0)   'Right hand side
            'These lines go from the top of the page to the bottom but Visual Basic can still work out what's
            'in the middle if you tell it.  Try changing the "280" to 75 and see what happens.
            'and then try changing the "174" to 95 and see what happens.
            'Make ONE change and print, then make the other change and print, that way there will be no confusion as to what change did what.
            
            
            ' Print variables in rows and then in columns.
              Printer.CurrentY = VerticalMargin + 23
             Printer.CurrentX = HorizontalMargin + 4
            
            Printer.Print "Print in Rows"
            Printer.CurrentY = VerticalMargin + 30
            '*********************************************************************
            'This section defines where the Columns will be
            'and then prints directly from the Array with a For Next loop "f"
            Col(0) = 4
            Col(1) = 52
            Col(2) = 100        'column width of 48 mm
            Col(3) = 148
            NR = 35             'starts 35 mm down from top boarder
             
            For f = LBound(AnArray) To UBound(AnArray)
                If AnArray(f) = Empty Then GoTo Skip     'Check for errors or empty's
                
              
                
             Printer.CurrentX = HorizontalMargin + (Col(k))  'eg:HorizontalMargin + (Column(2)) = 100 mm
            Printer.CurrentY = VerticalMargin + (NR)
            
            Printer.Print AnArray(f)
            k = k + 1               'Next Column on the next loop
            If k = 4 Then NR = NR + 7: k = 0       'When you reach Column 4, start the next Row at the first column
            
            
            'If NR > 270 Then Printer.NewPage: NR = 20  '****If you had to fill a page****
            
Skip:
            Next f
                        
                        
                        
                        Printer.DrawWidth = 4
                        Printer.Line (70 + HorizontalMargin, 105 + VerticalMargin)-(HorizontalMargin + 102, 182 + VerticalMargin), RGB(0, 0, 0), B
                         '^^^ This will draw a box around the centre column.
                         'Notice I set the DrawWidth to 4 and then set it back to 1. This is a good habit to get into as the settings
                         'once set, remain current until changed and can be accumulative.
                         'If your having trouble with positioning, check that your not adding onto
                         'a previous number
                         'you can see this in the above line. "Printer.Line (70 +" If you get
                         'a ruler you will see that 70 mm measures from the line we put down the Left hand side
                         'and not the edge or the 5 mm Margin we set earlier.
                         
                        
                         'Try commenting out the above line and uncommenting the next 3 lines and see what happens
                     '   HorizontalMargin = (210 - Printer.ScaleWidth) / 2       'Here the Margin is reset to the printer default.
                     '   Printer.Line (75 + HorizontalMargin, 105 + VerticalMargin)-(HorizontalMargin + 110, 182 + VerticalMargin), RGB(0, 0, 0), B
                     '   HorizontalMargin = 5 + HorizontalMargin     ' It is now reset to the + 5
                        
                        
                        Printer.DrawWidth = 1
            
            
            '************************************************************************************
                        'This section defines where the Columns will be
            'and then prints from the List Box with a For Next loop "Listword"
 
             
             
             M = 110
            
             k2 = 0
            Col2(0) = 10
            Col2(1) = 74
            Col2(2) = 136        'column width of 63 mm
            
            NC = M             'start at 83 down 10 across
              Printer.CurrentY = VerticalMargin + 100
             Printer.CurrentX = HorizontalMargin + 4
            
            Printer.Print "Print in Column's"
                        Printer.FontBold = False         'we don't want bold   ***** NOTE : We turned it off

            For ListWord = 0 To List1.ListCount - 1
            
            Printer.CurrentX = HorizontalMargin + (Col2(k2))
            Printer.CurrentY = VerticalMargin + (NC)
            Printer.Print List1.List(ListWord)
            
           NC = NC + 7


      If NC >= 180 Then k2 = k2 + 1: NC = M      'set the lower limit of the column
                                                                     ' "Then"  go to next column (k2)
                                                                     ' go to top of column (NC = M)
                        
   '***********************************************************************************************************
     'If k2 = 3 Then M = M + 70: k2 = 0                 'This is just to give you an idea of how
     'If M >= 250 Then Printer.NewPage                 ' to carry on to the next page if you had more
                                                                        ' data than what we are dealing with here.
 '*************************************************************************************************************
 Next
            
         Printer.PaintPicture Picture1.Picture, HorizontalMargin + 30, VerticalMargin + 5, 30, 18
         
         Printer.PaintPicture Picture1.Picture, HorizontalMargin + 115, VerticalMargin + 200, 50, 30
            
            
            
            
            
            Printer.FillStyle = 0   'Solid
          'object.Circle [Step] (x, y), radius, [color, start, end, aspect
            Printer.Circle (60, 150), 5     'Dot
            
            Printer.FillStyle = 1   '(Default) Transparent again
            
            
                        Printer.DrawWidth = 4
                        Printer.Circle (130, 150), 15

            
            
                      Printer.CurrentX = HorizontalMargin + 10
                      Printer.CurrentY = VerticalMargin + 223
                Printer.Print "Notes:"
                Printer.Line (HorizontalMargin + 30, VerticalMargin + 229)-(110 - HorizontalMargin, VerticalMargin + 229) ' WritingSpace
                Printer.Line (HorizontalMargin + 10, VerticalMargin + 237)-(110 - HorizontalMargin, VerticalMargin + 237) ' WritingSpace
                Printer.Line (HorizontalMargin + 10, VerticalMargin + 245)-(110 - HorizontalMargin, VerticalMargin + 245) ' WritingSpace
                Printer.Line (HorizontalMargin + 10, VerticalMargin + 253)-(110 - HorizontalMargin, VerticalMargin + 253) ' WritingSpace
                Printer.Line (HorizontalMargin + 10, VerticalMargin + 261)-(110 - HorizontalMargin, VerticalMargin + 261) ' WritingSpace
      
           Printer.CurrentX = 120
           Printer.CurrentY = 260
           Printer.Print Text1.Text
           
            
           Printer.CurrentX = 120
           Printer.CurrentY = 20
           Printer.Print Text2.Text
           
           

                       Printer.Line (HorizontalMargin + 117, VerticalMargin + 188)-(178 - HorizontalMargin, VerticalMargin + 196), RGB(0, 0, 0), B

           Printer.CurrentX = 131
           Printer.CurrentY = 201
           Printer.Print Text3.Text
           
            
            
            
            'Now, we'll centre a line of text in the page. To get the correct text measurements,
            
            'We put the text in a variable and get the text width
            CenteredText = "If you like this tuitorial then please vote."
            CenteredTextWidth = Printer.TextWidth(CenteredText)
            
            'To set a starting position, we use Printer.CurrentX and Printer.CurrentY
            'functions. To know where the text is to be located horizontally, we will use
            'a very simple formula:
            '            (Page Width - Text Width) / 2
            'For height, we will put it 220mm under the top of our printable area (towards the end of the page)
            Printer.CurrentX = (210 - CenteredTextWidth) / 2
            Printer.CurrentY = VerticalMargin + 265
            Printer.Print CenteredText
            Label1.Caption = "Finished"
            
            Label2.Caption = Line1 & Chr(10) & Line2 & Chr(10) & Line3 & Chr(10) & Line4 & Chr(10) & Line5
            Command1.Caption = "Credits"
            
            
            Printer.EndDoc
            TimesClicked = TimesClicked + 1
            Exit Sub
Cancel:                     ' The user has cancelled so:
            Text1.Visible = False
            Text2.Visible = False                 'Hide Stuff
            Text3.Visible = False
             List1.Visible = False

            Line1 = "You have decided to cancel this tuitorial."
            Line2 = "You can still check out the code and print at"
            Line3 = "a later time.    Click Next to Exit "


            Label1.Caption = "Finished"
            Label2.Caption = Line1 & Chr(10) & Line2 & Chr(10) & Line3 & Chr(10) & Line4 & Chr(10) & Line5

            Command1.Caption = "Next"

            TimesClicked = TimesClicked + 1
        
        Case Is = 5
        
            Line1 = "Thankyou for your time."
            Line2 = "Please go back to PSC and vote or"
            Line3 = "leave a comment, or both."
            Line4 = "My Site : Redback Studios"
            Line5 = "www.redbackstudios.com.au"
            Line6 = "Software Copyright 2001 by Digit Software. "
            Label1.Caption = "Credits"
            Command1.Caption = "Exit"
            Label2.Caption = Line1 & Chr(10) & Line2 & Chr(10) & Line3 & Chr(10) & Line4 & Chr(10) & Line5 & Chr(10) & Line6
            
            
            TimesClicked = TimesClicked + 1
        Case Is = 6
            Unload Me
            End




End Select

End Sub

Private Sub Form_Load()
Dim Counter As Integer
'TimesClicked = 0
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
 List1.Visible = False
  Picture1.Visible = False

Set Picture1.Picture = LoadPicture(App.Path & "\digit.emf")

'A For Next loop to fill the array we will use for printing "Rows and Columns"
For Counter = 0 To 29            '30 parts
AnArray(Counter) = "Any Text" & " " & (Counter)  ' eg: Any Text 25
List1.AddItem AnArray(Counter)
Next Counter

Command1.Caption = "Start"
Label1.Caption = "Click Start"
Label2.Caption = "And follow the on screen instructions."
End Sub


