VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Module Level Comment-----------------------------------------------------------
'
'Project name   : XMLTree
'Module name    : AddProperty
'Classification : Class Module
'Created        : 02 January 2001
'Author         : Paul Stevens
'Description    :
'
'Dependencies   : MSXML3, MSComctlLib.Treeview
'Issues         :
'Revision(s)    : Paul Stevens
'-------------------------------------------------------------------------------
Private Instance As Long

Public Function HasPropertys(Node As IXMLDOMNode, Treeview As Object, PropertyParentName As String)
Dim Propertys As IXMLDOMNode
Dim AddNode As New AddNode
On Error Resume Next
Set Propertys = Node
Dim Counter As Long
    'this will cycle through all of the nodes propertys and add them to the tree
    'EX: <File Name="Test"> Name is a Property of File with a Value of "Test"
    For Counter = 0 To Node.Attributes.length - 1
        AddNode.AddChildNode Propertys, CurrentNode, Treeview, True, Propertys.Attributes.Item(Counter).Text, Instance
    Next Counter
ParsingPropertys = False
End Function

Private Sub Class_Initialize()
    Instance = GlobalInstance + 1
    GlobalInstance = Instance
    Debug.Print Instance
End Sub
Private Sub Class_Terminate()
    GlobalInstance = GlobalInstance - 1
End Sub
