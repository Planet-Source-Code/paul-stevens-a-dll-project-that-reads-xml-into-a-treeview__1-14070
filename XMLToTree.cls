VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLToTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Module Level Comment-----------------------------------------------------------
'
'Project name   : XMLTree
'Module name    : XMLToTree
'Classification : Class Module
'Created        : 02 January 2001
'Author         : Paul Stevens
'Description    :
'
'Dependencies   : MSXML3, MSComctlLib.Treeview
'Issues         :
'Revision(s)    : Paul Stevens
'-------------------------------------------------------------------------------
Public Event NodeChange(NodeName As String, Nodenumber As Long, TotalNodes As Long)


Public Function PlantTree(Treeview As Object, Optional XMLFile As String, Optional XML As String, Optional File As Boolean = True)
Dim NewTree As MSComctlLib.Treeview
Dim XMLReader As MSXML2.DSOControl30
Dim XMLDocument As MSXML2.DOMDocument30
Dim Children As Long
Dim XMLRootNode As IXMLDOMNode
Dim XMLNode As IXMLDOMNode
Dim XMLTopNode As IXMLDOMNode
Dim Counter As Long
Dim Counter1 As Long

On Error Resume Next

Set XMLReader = New MSXML2.DSOControl30
Set NewTree = Treeview
If File Then
    XMLReader.XMLDocument.Load (XMLFile)
Else
    XMLReader.XMLDocument.loadXML (XML)
End If

Set XMLDocument = XMLReader.XMLDocument
Set XMLReader = New MSXML2.DSOControl30

RaiseEvent NodeChange("start", "0", XMLDocument.childNodes.length)

Set NewTree = Treeview
For Counter = 0 To XMLDocument.childNodes.length - 1
    Set XMLRootNode = XMLDocument.childNodes(Counter)
    If XMLRootNode.nodeType = NODE_ELEMENT Then
        NewTree.Nodes.Add , , XMLRootNode.NodeName, XMLRootNode.NodeName
    End If
        
            If XMLRootNode.hasChildNodes Then
                For Counter1 = 0 To XMLRootNode.childNodes.length - 1
                    Set XMLNode = XMLRootNode.childNodes(Counter1)
                    AddRootChild XMLNode, XMLNode.ParentNode.NodeName, Treeview
                Next Counter1
            Else
              
                    Set XMLNode = XMLRootNode
                    AddRootChild XMLNode, XMLNode.ParentNode.NodeName, Treeview

            End If
Next Counter
    
End Function

Private Function AddRootChild(NodeName As IXMLDOMNode, ParentNode As String, Treeview As Object, Optional IsProperty As Boolean = False, Optional Property As String)
    Dim AddNode As New AddNode
    Dim AddChild As AddChild
    Dim AddProperty As AddProperty
    On Error Resume Next
    Set AddChild = New AddChild
    Set AddProperty = New AddProperty
    
        isRootChild = True
        CurrentNode = NodeName.NodeName
        RaiseEvent NodeChange("start", "0", NodeName.childNodes.length)
        AddNode.AddChildNode NodeName, ParentNode, Treeview
        isRootChild = False
    If Not NodeName.nodeTypeString = "comment" Then
        If NodeName.Attributes.length > 0 Then
            If Not ParsingPropertys Then
            ParsingPropertys = True
            AddProperty.HasPropertys NodeName, Treeview, CurrentNode
            End If
        End If
    End If
        
        If NodeName.hasChildNodes Then
            AddChild.HasChild NodeName, Treeview, CurrentParentNode(0)
        End If
    
        
End Function


