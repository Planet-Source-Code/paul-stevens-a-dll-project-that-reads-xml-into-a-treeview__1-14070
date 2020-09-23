Attribute VB_Name = "Global"
'Module Level Comment-----------------------------------------------------------
'
'Project name   : XMLTree
'Module name    : Global
'Classification : Class Module
'Created        : 02 January 2001
'Author         : Paul Stevens
'Description    : Used To Store Global Variables
'
'Dependencies   :
'Issues         :
'Revision(s)    : Paul Stevens
'-------------------------------------------------------------------------------
Public NodeCount As Long
Public CurrentNode As String
Public CurrentParentNode(999) As String
Public ParsingPropertys As Boolean
Public isRootChild As Boolean
Public isParentNode As Boolean
Public ChildInstance As Long
Public NodeInstance As Long
Public PropertyInstance As Long
Public GlobalInstance As Long
