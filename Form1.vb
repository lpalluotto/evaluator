Public Class Form1
    Inherits System.Windows.Forms.Form

   WithEvents Evaluator1 As Evaluator

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Button1 As System.Windows.Forms.Button
   Friend WithEvents tbExpression As System.Windows.Forms.TextBox
   Friend WithEvents cbSamples As System.Windows.Forms.ComboBox
   Friend WithEvents tbResults As System.Windows.Forms.TextBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbExpression = New System.Windows.Forms.TextBox()
        Me.cbSamples = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbResults = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'tbExpression
        '
        Me.tbExpression.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbExpression.Location = New System.Drawing.Point(134, 37)
        Me.tbExpression.Name = "tbExpression"
        Me.tbExpression.Size = New System.Drawing.Size(168, 22)
        Me.tbExpression.TabIndex = 0
        '
        'cbSamples
        '
        Me.cbSamples.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSamples.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbSamples.Items.AddRange(New Object() {"<enter an expression or select a sample>", "123+345", "(2+3)*(4-2)", "120+5%", "now", "Round(now - Date(2005,1,1))+"" days since new year""", "anumber*5", "theForm.Left"})
        Me.cbSamples.Location = New System.Drawing.Point(134, 9)
        Me.cbSamples.Name = "cbSamples"
        Me.cbSamples.Size = New System.Drawing.Size(264, 24)
        Me.cbSamples.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(10, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 23)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Expression"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbResults
        '
        Me.tbResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbResults.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbResults.Location = New System.Drawing.Point(10, 92)
        Me.tbResults.Multiline = True
        Me.tbResults.Name = "tbResults"
        Me.tbResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.tbResults.Size = New System.Drawing.Size(388, 227)
        Me.tbResults.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(10, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Samples"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(10, 65)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 23)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Results"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(312, 37)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(90, 26)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Evaluate"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(408, 326)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.tbResults)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbSamples)
        Me.Controls.Add(Me.tbExpression)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Name = "Form1"
        Me.Text = "Expression Evaluator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
      Try
            AddLine(tbExpression.Text)
            Dim value As String = tbExpression.Text.Replace(",", "")
            Dim res As String = Evaluator1.Eval(value).ToString
            AddLine(" => " & res)
      Catch ex As Exception
         AddLine(" => Error :" & ex.Message)
      End Try
   End Sub

   Private Sub AddLine(ByVal s As String)
      tbResults.AppendText(s & vbCrLf)
      tbResults.SelectionStart = tbResults.Text.Length
   End Sub

   Private Sub cbSamples_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSamples.SelectedIndexChanged
      tbExpression.Text = cbSamples.Text
   End Sub

   Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Evaluator1 = New Evaluator
      cbSamples.SelectedIndex = 0
   End Sub


   Private Sub Evaluator1_GetVariable(ByVal name As String, ByRef value As Object) Handles Evaluator1.GetVariable
      Select Case name
         Case "anumber"
            value = 5.0
         Case "adate"
            value = #1/1/2005#
         Case "theForm"
            value = Me
      End Select
   End Sub
End Class
