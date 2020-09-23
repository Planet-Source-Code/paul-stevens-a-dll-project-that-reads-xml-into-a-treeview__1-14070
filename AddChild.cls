VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Module Level Comment-----------------------------------------------------------
'
'Project name   : XMLTree
'Module name    : AddChild
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
Public Function HasChild(Node As IXMLDOMNode, Treeview As Object, ChildParentName As String)
Dim Counter As Long
Dim AddNode As New AddNode
Dim AddProperty As New AddProperty

On Error Resume Next

    'we will loop through each of the child nodes of the current parent node
    For Counter = 0 To Node.childNodes.length - 1
        AddNode.AddChildNode Node.childNodes(Counter), ChildParentName, Treeview, , , Instance
        'check if the node has propertys inside of it and add them as neccesary
        If Node.Attributes.length > 0 Then
            If Not ParsingPropertys Then
                ParsingPropertys = True
                AddProperty.HasPropertys Node, Treeview, CurrentNode
            End If
        End If
    Next Counter

End Function

Private Sub Class_Initialize()
    Instance = GlobalInstance + 1
    GlobalInstance = Instance
    Debug.Print Instance
End Sub

Private Sub Class_Terminate()
    GlobalInstance = GlobalInstance - 1
End Sub
