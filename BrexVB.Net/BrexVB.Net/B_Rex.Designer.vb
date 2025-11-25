<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class B_Rex
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    ' ComboBox-Feld für Templates (wird zur Laufzeit nach ReadData erstellt)
    Private WithEvents cmbTemplates As System.Windows.Forms.ComboBox

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        StartCode.ReadData()

        ' ComboBox erst nach dem Einlesen der Daten erstellen und füllen
        cmbTemplates = New System.Windows.Forms.ComboBox()
        With cmbTemplates
            .Name = "cmbTemplates"
            .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            .Dock = System.Windows.Forms.DockStyle.Top
            .TabIndex = 0
        End With

        ' Direkter Zugriff auf StartCode.Templates (Index 1..12)
        For i As Integer = 1 To Math.Min(12, StartCode.Templates.Length - 1)
            Dim t = StartCode.Templates(i)
            If t IsNot Nothing AndAlso Not String.IsNullOrEmpty(t.Bezeichnung) Then
                cmbTemplates.Items.Add(t.Bezeichnung)
            End If
        Next

        ' ComboBox zur Form hinzufügen
        Me.Controls.Add(cmbTemplates)
    End Sub

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ' Form-Größe verdoppelt
        Me.ClientSize = New System.Drawing.Size(1600, 900)
        Me.Text = "B_Rex"

        ' Daten einlesen
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    ' Handler für Auswahlwechsel / "Click" eines Combobox-Items
    Private Sub cmbTemplates_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTemplates.SelectedIndexChanged
        Dim cb = DirectCast(sender, System.Windows.Forms.ComboBox)
        If cb.SelectedIndex >= 0 Then
            Dim selectedText = cb.SelectedItem?.ToString()
            System.Diagnostics.Debug.WriteLine($"Template ausgewählt: {selectedText}")
            ' Beispiel: Zugriff auf Template-Objekt, falls weitere Daten benötigt werden
            Dim idx = cb.SelectedIndex + 1 ' Items starten bei 0, Templates bei 1
            If idx >= 1 AndAlso idx <= 12 Then
                Dim tmpl = StartCode.Templates(idx)
                If tmpl IsNot Nothing Then
                    ' TODO: Aktion mit tmpl.Anlage oder anderen Feldern
                    System.Diagnostics.Debug.WriteLine($"Anlage-Inhalt Länge: {If(tmpl.Anlage, String.Empty).Length}")
                End If
            End If
        End If
    End Sub

End Class
