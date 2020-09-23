VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Module Level Comment-----------------------------------------------------------
'
'Project name   : XMLTree
'Module name    : AddNode
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
Public Function AddChildNode(NodeName As IXMLDOMNode, ParentNode As String, Treeview As Object, Optional IsProperty As Boolean = False, Optional Property As String, Optional AddInstance As Long)
Dim ParentTree As MSComctlLib.Treeview
Dim ParentName As String
Dim AddChild As AddChild

Set ParentTree = Treeview

'Because we can not determine if a node already exists in the tree
'we will try to add it and if it does not work handle the error
On Error GoTo ErrHandler
ParentName = ""
'First Make sure that this is not a root Node
If Not ParentNode = "#document" Then
        'Keep track of how many nodes we have
        NodeCount = NodeCount + 1
    ' if we are not adding a property
    If Not IsProperty Then
        'Make sure that what we are adding is really a Node Element
        If NodeName.nodeType = NODE_ELEMENT Then
                AddToTree ParentTree, ParentNode, tvwChild, NodeName.NodeName, NodeName.NodeName
                'Set the current node to the one that was just added
                CurrentNode = NodeName.NodeName
                'Set the ParentNode for this instance of the class adding the node
                'To the Node that was just added (AddChild Or AddProperty Class)
                CurrentParentNode(AddInstance) = NodeName.NodeName
        ElseIf NodeName.nodeType = NODE_TEXT Then
            'If the Node is a Value not an elament add the value
            If Not IsNumeric(NodeName.nodeValue) Then
                AddToTree ParentTree, ParentNode, tvwChild, NodeName.nodeValue, NodeName.nodeValue
                CurrentNode = CurrentParentNode(AddInstance)
                CurrentParentNode(AddInstance) = NodeName.ParentNode.ParentNode.NodeName
            Else
                AddToTree ParentTree, ParentNode, tvwChild, "K" & NodeName.nodeValue, NodeName.nodeValue
                CurrentNode = CurrentParentNode(AddInstance)
                CurrentParentNode(AddInstance) = NodeName.ParentNode.ParentNode.NodeName
            End If
        Else
            'if its neither an Element or a Vlaue then we add the Raw XML
            AddToTree ParentTree, ParentNode, tvwChild, NodeName.XML, NodeName.XML
                CurrentNode = NodeName.XML
                CurrentParentNode(AddInstance) = NodeName.XML
        End If
    Else
    'This means we are adding a property
        If Not IsNumeric(Property) Then
            AddToTree ParentTree, ParentNode, tvwChild, Property, Property
            'We wont set the parent Node here as the Property Can not be  a Parent
            If isRootChild Then
                CurrentNode = Property
            End If
        Else
           AddToTree ParentTree, ParentNode, tvwChild, "P" & Property, Property
           If isRootChild Then
                CurrentNode = "P" & Property
            End If
        End If
    End If
End If
    If Not isRootChild Then
        If NodeName.hasChildNodes Then
            'Check if the Node has any Children, if it has instantiate a new
            'Instance of AddChild and add the Child Node
            Set AddChild = New AddChild
            AddChild.HasChild NodeName, Treeview, CurrentParentNode(AddInstance)
        End If
    End If
Exit Function
ErrHandler:
    'if an error occured the we check the error Number
    'If its = 35602 that means the key already exists
    'in the Tree and we need to use another key
    'the Code that follows is not very pretty but it
    'creates unique keys for each node it adds
    'other than that it is a complete replica of the code above
    If Err.Number = 35602 Then
        If Not IsProperty Then
             If NodeName.nodeType = NODE_ELEMENT Then
                 AddToTree ParentTree, ParentNode, tvwChild, ParentNode & NodeCount & NodeName.NodeName, NodeName.NodeName
            If isRootChild Then
                 CurrentNode = ParentNode & NodeCount & NodeName.NodeName
            End If
                 CurrentParentNode(AddInstance) = ParentNode & NodeCount & NodeName.NodeName
             Else
                 If Not IsNumeric(NodeName.nodeValue) Then
                    AddToTree ParentTree, ParentNode, tvwChild, ParentNode & NodeCount & NodeName.nodeValue, NodeName.nodeValue
                    CurrentParentNode(AddInstance) = ParentNode & NodeCount & NodeName.nodeValue
                    CurrentNode = CurrentParentNode(AddInstance)
                    CurrentParentNode(AddInstance) = NodeName.ParentNode.ParentNode.NodeName
                 Else
                    AddToTree ParentTree, ParentNode, tvwChild, "K" & NodeCount & NodeName.nodeValue, NodeName.nodeValue
                    CurrentParentNode(AddInstance) = "K" & NodeCount & NodeName.nodeValue
                    CurrentNode = CurrentParentNode(AddInstance)
                    CurrentParentNode(AddInstance) = NodeName.ParentNode.ParentNode.NodeName
                 End If
             End If
        Else
            AddToTree ParentTree, ParentNode, tvwChild, ParentNode & NodeCount & Property, Property
            CurrentParentNode(AddInstance) = ParentNode & NodeCount & Property
            If isRootChild Then
                CurrentNode = ParentNode & NodeCount & Property
            End If
        End If
    If Not isRootChild Then
        If NodeName.hasChildNodes Then
            Set AddChild = New AddChild
            AddChild.HasChild NodeName, Treeview, CurrentParentNode(AddInstance)
        End If
    End If
    ElseIf Err.Number = 424 Then
        Resume Next
    End If
End Function
Private Sub Class_Initialize()
    Instance = GlobalInstance + 1
    GlobalInstance = Instance
    Debug.Print Instance
End Sub
Private Sub Class_Terminate()
    GlobalInstance = GlobalInstance - 1
End Sub
Private Function AddToTree(Tree As Object, Optional Relative, Optional Relationship, Optional Key, Optional Text, Optional Image, Optional SelectedImage)
Dim NewNode As MSComctlLib.Treeview
Set NewNode = Tree
'This adds the node to the tree
NewNode.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage

End Function
