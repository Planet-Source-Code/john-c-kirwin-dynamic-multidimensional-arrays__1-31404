VERSION 5.00
Begin VB.Form frm2DArray 
   Caption         =   "Dynamic Multidimensional Array"
   ClientHeight    =   1755
   ClientLeft      =   12750
   ClientTop       =   2745
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display"
      Default         =   -1  'True
      Height          =   350
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the button below to display the elements of a two-dimensional array"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frm2DArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//****************************************************************************
'// Dynamic Multidimensional Arrays - Simple example of a two-dimensional array.
'// It has rows and columns just like a spreadsheet in a program like MS Excel.
'// The array in this demo is dynamic because it is redimmable.
'// by John Kirwin (02/01/02)
'//****************************************************************************

'//****  The Option Base statement makes code reuse more difficult, especially
'//      if you like to cut and paste, because you have to be aware of what the
'//      code base was originally. The Option Base 1 sets the beginning indices
'//      at 1 instead of 0 the default.
Option Base 1

'//****************************************************************************
'// Global Variables
'//****************************************************************************
Dim iRows As Integer                                                           ' Total Rows in the Array
Dim iColumns As Integer                                                        ' Total Columns in the Array

'//**** Declare Dynamic Array
Dim MyArray() As String
Private Sub Form_Load()
'//****************************************************************************
'// Form_Load - In the form load event the dynamic array MyArray() is
'// redimensioned and the demonstration elements are loaded.
'//****************************************************************************
On Error GoTo EH

    '//**** Define the size of the dynamic array MyArray()
    iRows = 2
    iColumns = 3

    '//**** Re Dimension the two-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim MyArray(iRows, iColumns)

    '//**** Load Row 1
    MyArray(1, 1) = "A1"
    MyArray(1, 2) = "B1"
    MyArray(1, 3) = "C1"

    '//**** Load Row 2
    MyArray(2, 1) = "A2"
    MyArray(2, 2) = "B2"
    MyArray(2, 3) = "C2"

    '//**** Exit Sub/Function before error handler
    Exit Sub

EH:

    '//**** Error Handling
    MsgBox Err.Number & ": " & Err.Description & _
           " occurred during the Form_Load" _
           , vbInformation, " Warning"

End Sub
Private Sub Command1_Click()
'//****************************************************************************
'// Command1_Click - Add a button Command1 and use the click event for
'// displaying results
'//****************************************************************************
On Error GoTo EH
   
    '//**** Variable Declarations
    Dim Row As Integer                                                         ' Track Row number in routine
    Dim Column As Integer                                                      ' Track Column number in routine
    Dim sRow As String                                                         ' String to add row to listbox

    '//**** For Next through the range of indices
    '//     base 1 to the total rows iRow
    For Row = 1 To iRows
        
        '//**** For Next through the range of indices
        '//     base 1 to the total columns iColumns
        For Column = 1 To iColumns
        
            '//**** Add each column element to the string
            '//     sRow used to add row to listbox
            sRow = sRow & Space$(3) & MyArray(Row, Column)
        
        '//**** Increment Column For/Next Loop
        Next Column
        
        '//**** Using listbox's AddItem Method add the
        '//     string sRow to listbox display of the
        '//     elements of the array
        Call List1.AddItem(sRow)
   
        '//**** Clear sRow string that is used to add row to listbox
        sRow = ""
    
    '//**** Increment Row For/Next Loop
    Next Row

    
    '//**** Change instruction label caption to
    '//     instruct users to check out the code
    Label1.Caption = "Now go inspect the code to see how easy it is!"
    
    '//**** Hide Display button
    Command1.Visible = False
    Command1.Enabled = False
    
    '//**** Display Close button
    Command2.Visible = True
    Command2.Enabled = True
    
    '//**** Exit Sub/Function before error handler
    Exit Sub

EH:

    '//**** Error Handling
    MsgBox Err.Number & ": " & Err.Description & _
           " occurred during the Command1_Click" _
           , vbInformation, " Warning"

End Sub
Private Sub Command2_Click()
'//****************************************************************************
'// Cancel Button Click Event
'//****************************************************************************
On Error Resume Next
    
    '//**** Unload the form
    Unload frm2DArray
     
    '//**** Set the form to nothing
    Set frm2DArray = Nothing
    
    '//**** End the application
    End
    
End Sub


