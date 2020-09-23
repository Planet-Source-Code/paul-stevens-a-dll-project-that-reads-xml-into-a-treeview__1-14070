Option Strict Off
Option Explicit Off
Namespace Project2
	Public  Class Form1
		Inherits System.WinForms.Form
		Private Shared  m_vb6FormDefInstance As Form1
		Public Shared  Property DefInstance() As Form1
			Get
				If m_vb6FormDefInstance Is Nothing Then
					m_vb6FormDefInstance = New Form1()
				End If
				DefInstance = m_vb6FormDefInstance
			End Get
			Set
				m_vb6FormDefInstance = Value
			End Set
		End Property
#Region " Windows Form Designer generated code "
		' Required by the Win Form Designer
		Private  components As System.ComponentModel.Container
		Public  ToolTip1 As System.WinForms.ToolTip
		Public  Tag1 As Microsoft.VisualBasic.Compatibility.VB6.Tag
		Private WithEvents  Form1 As Form1
		Public WithEvents  pb As AxMSComctlLib.AxProgressBar
		Public WithEvents  Command2 As System.WinForms.Button
		Public WithEvents  d1 As AxMSComDlg.AxCommonDialog
		Public WithEvents  Command1 As System.WinForms.Button
		Public WithEvents  TreeView1 As AxMSComctlLib.AxTreeView
		Public Sub New()
			MyBase.New()
			Me.Form1 = Me
			' This call is required by the Win Form Designer
			If m_vb6FormDefInstance Is Nothing Then
				m_vb6FormDefInstance = Me
			End If
			InitializeComponent()
		End Sub
		' Form overrides dispose to clean up the component list.
		Public Overrides Sub Dispose()
			MyBase.Dispose()
			components.Dispose()
		End Sub
		' The main entry point for the application
		Shared Sub Main()
			System.WinForms.Application.Run(New Form1())
		End Sub
		' NOTE: The following procedure is required by the Win Form Designer
		' It can be modified using the Win Form Designer.
		' Do not modify it using the code editor.

		Private Sub InitializeComponent()
			Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager("Form1", GetType(Form1), Nothing, True)
			Me.components = New System.ComponentModel.Container()
			Me.ToolTip1 = New System.WinForms.ToolTip(components)
			Me.Tag1 = New Microsoft.VisualBasic.Compatibility.VB6.Tag()
			ToolTip1.Active = True
			' @design ToolTip1.SetLocation(New System.Drawing.Point(7,7))
			' @design Tag1.SetLocation(New System.Drawing.Point(90,7))
			Me.pb = New AxMSComctlLib.AxProgressBar
			Me.Command2 = New System.WinForms.Button
			Me.d1 = New AxMSComDlg.AxCommonDialog
			Me.Command1 = New System.WinForms.Button
			Me.TreeView1 = New AxMSComctlLib.AxTreeView
			pb.BeginInit()
			d1.BeginInit()
			TreeView1.BeginInit()
			Me.ControlBox = True
			Me.Text = "Form1"
			Me.ClientSize = New System.Drawing.Size(534, 518)
			Me.Location = New System.Drawing.Point(4, 23)
			Me.StartPosition = System.WinForms.FormStartPosition.WindowsDefaultLocation
			Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
			Me.BackColor = System.Drawing.SystemColors.Control
			Me.BorderStyle = System.WinForms.FormBorderStyle.Sizable
			Me.Enabled = True
			Me.KeyPreview = False
			Me.MaximizeBox = True
			Me.MinimizeBox = True
			Me.Cursor = System.Drawing.Cursors.Default
			Me.RightToLeft = System.WinForms.RightToLeft.No
			Me.ShowInTaskbar = True
			Me.HelpButton = False
			Me.WindowState = System.WinForms.FormWindowState.Normal
			pb.OcxState = CType(resources.GetObject("pb.OcxState"), System.WinForms.AxHost.State)
			pb.Size = New System.Drawing.Size(169, 17)
			pb.Location = New System.Drawing.Point(360, 104)
			pb.TabIndex = 3
			Command2.Text = "Command2"
			Command2.Size = New System.Drawing.Size(81, 33)
			Command2.Location = New System.Drawing.Point(376, 48)
			Command2.TabIndex = 2
			Command2.FlatStyle = System.WinForms.FlatStyle.Standard
			Command2.BackColor = System.Drawing.SystemColors.Control
			Command2.CausesValidation = True
			Command2.Enabled = True
			Command2.Cursor = System.Drawing.Cursors.Default
			Command2.RightToLeft = System.WinForms.RightToLeft.No
			Command2.TabStop = True
			d1.OcxState = CType(resources.GetObject("d1.OcxState"), System.WinForms.AxHost.State)
			d1.Location = New System.Drawing.Point(496, 8)
			Command1.Text = "Command1"
			Command1.Size = New System.Drawing.Size(81, 33)
			Command1.Location = New System.Drawing.Point(376, 8)
			Command1.TabIndex = 1
			Command1.FlatStyle = System.WinForms.FlatStyle.Standard
			Command1.BackColor = System.Drawing.SystemColors.Control
			Command1.CausesValidation = True
			Command1.Enabled = True
			Command1.Cursor = System.Drawing.Cursors.Default
			Command1.RightToLeft = System.WinForms.RightToLeft.No
			Command1.TabStop = True
			TreeView1.OcxState = CType(resources.GetObject("TreeView1.OcxState"), System.WinForms.AxHost.State)
			TreeView1.Size = New System.Drawing.Size(337, 505)
			TreeView1.Location = New System.Drawing.Point(16, 0)
			TreeView1.TabIndex = 0
			Me.Controls.Add(pb)
			Me.Controls.Add(Command2)
			Me.Controls.Add(d1)
			Me.Controls.Add(Command1)
			Me.Controls.Add(TreeView1)
			pb.EndInit()
			d1.EndInit()
			TreeView1.EndInit()
		End Sub
#End Region 
		Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
			Dim x As XMLTree.XMLToTree
			On Error Goto ErrHandler
			x = New XMLTree.XMLToTree
			d1.CancelError = True
			d1.ShowOpen()
			x.PlantTree(CStr(VB6.GetDefaultProperty(TreeView1)), d1.FileName)
ErrHandler: 
		End Sub
		
		Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
			TreeView1.Nodes.Clear()
		End Sub
	End Class
End NameSpace