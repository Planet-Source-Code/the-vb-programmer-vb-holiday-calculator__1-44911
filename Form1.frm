VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   1380
   ClientTop       =   1410
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9165
   Begin VB.CommandButton Command1 
      Caption         =   "Show Holidays"
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
      Left            =   180
      TabIndex        =   0
      Top             =   6540
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' VB holiday calculator
' --- Bruce Gordon 4/20/03
' Use to get the dates for the most common American holidays
' as well as the major Jewish holidays.
' It was easy enough for me to code the American holidays
' (except for Easter, but the algorithm for that can be found
' on many sites). After a fairly lengthy search, I found usable
' code for the Jewish holidays on this gentleman's site:
'   http://www.geocities.com/couprie/calmath/
' (where I unabashedly lifted the desired code). Mr. Couprie also
' has logic for Islamic and Persian events as well.


Private Sub Command1_Click()

    Dim intYear     As Integer
    Dim strHoliday  As String
    Dim dtmTestDate As Date

   ' assume the user puts in a valid year - no error checking here
    intYear = Val(InputBox("Please enter desired year:"))

    Cls
    Print "Holidays for " & intYear
    dtmTestDate = DateSerial(intYear, 1, 1)
    Do While dtmTestDate <= DateSerial(intYear, 12, 31)
        strHoliday = GetHolidayName(dtmTestDate)
        If strHoliday <> "" Then
            Print MonthName(Month(dtmTestDate), True) & " " _
                & Day(dtmTestDate) & ":" _
                & vbTab & strHoliday
        End If
        dtmTestDate = DateAdd("d", 1, dtmTestDate)
    Loop

End Sub
