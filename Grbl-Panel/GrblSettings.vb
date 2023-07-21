Imports System.Windows.Forms
Imports System.Threading.Thread
Imports GrblPanel.My.Resources
Imports System.IO


Partial Class GrblGui
    Public Class GrblSettings
        ' Handle settings related operations


        Private _gui As GrblGui
        Private _paramTable As DataTable    ' to hold the parameters
        Private _nextParam As Integer       ' to track which param line is next

        Private _settings As Dictionary(Of String, String) = New Dictionary(Of String, String) From {
            {"0", Resources.GrblSettings_StepPulseUsec},
            {"1", Resources.GrblSettings_StepIdleDelayMsec},
            {"2", Resources.GrblSettings_StepPortInvertMask},
            {"3", Resources.GrblSettings_DirPortInvertMask},
{"4", Resources.GrblSettings_StepEnableInvertBool},
            {"5", Resources.GrblSettings_LimitPinsInvertBool},
            {"6", Resources.GrblSettings_ProbePinInvertBool},
            {"10", Resources.GrblSettings_StatusReportMask},
            {"11", Resources.GrblSettings_JunctionDeviationMm},
            {"12", Resources.GrblSettings_ArcToleranceMm},
            {"13", Resources.GrblSettings_ReportInchesBool},
            {"20", Resources.GrblSettings_SoftLimitsBool},
            {"21", Resources.GrblSettings_HardLimitsBool},
            {"22", Resources.GrblSettings_HomingCycleBool},
            {"23", Resources.GrblSettings_HomingDirInvertMask},
            {"24", Resources.GrblSettings_HomingFeedMmMin},
            {"25", Resources.GrblSettings_HomingSeekMmMin},
            {"26", Resources.GrblSettings_HomingDebounceMsec},
            {"27", Resources.GrblSettings_HomingPullOffMm},
            {"30", Resources.GrblSettings_RpmMax},
            {"31", Resources.GrblSettings_RpmMin},
{"32", Resources.GrblSettings_LaserMode},
            {"100", Resources.GrblSettings_XStepMm},
            {"101", Resources.GrblSettings_YStepMm},
            {"102", Resources.GrblSettings_ZStepMm},
            {"103", Resources.GrblSettings_AStepMm},
            {"110", Resources.GrblSettings_XMaxRateMmMin},
            {"111", Resources.GrblSettings_YMaxRateMmMin},
            {"112", Resources.GrblSettings_ZMaxRateMmMin},
            {"113", Resources.GrblSettings_AMaxRateMmMin},
            {"120", Resources.GrblSettings_XAccelMmSec2},
            {"121", Resources.GrblSettings_YAccelMmSec2},
            {"122", Resources.GrblSettings_ZAccelMmSec2},
            {"123", Resources.GrblSettings_AAccelMmSec2},
            {"130", Resources.GrblSettings_XMaxTravelMm},
            {"131", Resources.GrblSettings_YMaxTravelMm},
            {"132", Resources.GrblSettings_ZMaxTravelMm},
            {"133", Resources.GrblSettings_AMaxTravelMm}
            }
#Region "Properties"
        ReadOnly Property IsHomingEnabled As Integer
            Get
                Dim row As DataRow
                row = _paramTable.Rows.Find("$22")
                If Not IsNothing(row) Then
                    Return row.Item(Resources.GrblSettings_FillSettings_Value)
                End If
                Return 0
            End Get
        End Property
        ''' <summary>
        ''' Gets a value indicating whether Grbl is override capable.
        ''' </summary>
        ''' <value>
        ''' <c>true</c> if Grbl is override capable; otherwise, <c>false</c>.
        ''' </value>
        ReadOnly Property IsOverrideCapable As Boolean
            Get
                Dim row As DataRow
                row = _paramTable.Rows.Find("$31")
                If Not IsNothing(row) Then
                    Return True
                End If
                Return False
            End Get
        End Property
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        ReadOnly Property IsGrblMetric As Boolean
            Get
                Dim row As DataRow
                row = _paramTable.Rows.Find("$13")
                If Not IsNothing(row) Then
                    If (row.Item(1).ToString(0) = "0") Then
                        Return True
                    End If
                End If
                Return False
            End Get
        End Property
#End Region

        Public Sub New(ByRef gui As GrblGui)
            ' Get ref to parent object
            _gui = gui
            ' For Connected events
            AddHandler(_gui.Connected), AddressOf GrblConnected

        End Sub
        Public Sub EnableState(ByVal yes As Boolean)

            If yes Then
                _gui.gbGrblSettings.Enabled = True
            Else
                _gui.gbGrblSettings.Enabled = False
            End If
        End Sub

        Private Sub GrblConnected(ByVal msg As String)     ' Handles GrblGui.Connected Event
            If msg = "Connected" Then

                ' We are connected to Grbl so populate the Settings
                _nextParam = 0
                gcode.sendGCodeLine("$$")
            End If
        End Sub

        Public Sub FillSettings(ByVal data As String)
            ' Add a settings line to the display
            'Console.WriteLine("GrblSettings: $ Data is : " + data)
            ' Return
            Dim params() As String
            Dim index As Integer
            If _nextParam = 0 Then
                ' We are dealing with a fresh
                _paramTable = New DataTable
                With _paramTable
                    .Columns.Add("ID")
                    .Columns.Add(Resources.GrblSettings_FillSettings_Value)
                    .Columns.Add("Description")
                    .PrimaryKey = New DataColumn() { .Columns("ID")}
                End With
            End If
            params = data.Split({"="c, "("c, ")"c})
            params(1) = params(1).Replace(" ", "")         ' strip trailing blanks
            If (params.Count = 4) Then
                _paramTable.Rows.Add(params(0), params(1), params(2))
            Else
                ' We have Grbl in GUI mode, so add Description manually
                index = params(0).Remove(0, 1)      ' remove the leading $
                If _settings.ContainsKey(index) Then
                    _paramTable.Rows.Add(params(0), params(1), _settings(index))
                Else
                    _paramTable.Rows.Add(params(0), params(1), "")
                End If
            End If
            _nextParam += 1
            If params(0) = _gui.tbSettingsGrblLastParam.Text Then ' We got the last one
                _nextParam = 0            ' in case user does a MDI $$
                With _gui.dgGrblSettings
                    .DataSource = _paramTable
                    .Columns("ID").Width = 40
                    .Columns("ID").ReadOnly = True
                    .Columns("ID").DefaultCellStyle.BackColor = SystemColors.Control
                    .Columns(Resources.GrblSettings_FillSettings_Value).Width = 60
                    .Columns("Description").Width = 200
                    .Columns("Description").ReadOnly = True
                    .Columns("Description").DefaultCellStyle.BackColor = SystemColors.Control
                    .Refresh()
                End With
                ' Tell everyone we have the params
                RaiseEvent GrblSettingsRetrievedEvent()
            End If
        End Sub

        ' Event template for Settings Retrieved indication
        Public Event GrblSettingsRetrievedEvent()

        Public Sub RefreshSettings()
            _nextParam = 0
            gcode.sendGCodeLine("$$")
        End Sub

        ''' <summary>
        ''' Exports the Grbl settings to a file with a .grbls extension.
        ''' The user is prompted to choose a location to save the file using a SaveFileDialog.
        ''' The settings are written to the file in the format "ID=Value".
        ''' </summary>
        ''' <remarks>
        ''' This function uses the _paramTable DataTable, which contains the Grbl settings.
        ''' </remarks>
        ''' <exception cref="System.Exception">Thrown if there is an error during the export process.</exception>
        ''' <param name="progressBar">A ProgressBar control to display the export progress. This can be null if not used.</param>
        Public Sub ExportSettingsToFile(progressBar As ProgressBar)
            ' Check if dgGrblSettings has rows
            If _gui.dgGrblSettings.Rows.Count = 0 Then
                MessageBox.Show("No settings to export.")
                Return
            End If
            ' Initialize saveFileDialog
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Grbl settings files (*.grbls)|*.grbls"
            ' Check if path is browsed correctly
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                progressBar.Visible = True
                Try
                    Dim filePath As String = saveFileDialog.FileName

                    ' Create a StreamWriter to write the settings data to the file
                    Using writer As New StreamWriter(filePath)
                        ' Loop through the _paramTable and write each setting as a line in the file
                        For Each row As DataRow In _paramTable.Rows
                            Dim paramLine As String = row("ID").ToString() & "=" & row(Resources.GrblSettings_FillSettings_Value).ToString()
                            writer.WriteLine(paramLine)
                        Next
                    End Using
                    progressBar.Value = 100
                    MessageBox.Show("Successfully exported settings!")
                    ' Reset the progress bar
                    progressBar.Value = 0
                    progressBar.Visible = False
                Catch ex As Exception
                    ' Reset the progress bar
                    progressBar.Value = 0
                    progressBar.Visible = False
                    MessageBox.Show("Failed to export settings : " & Environment.NewLine & ex.Message)
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Imports Grbl settings from a file with a .grbls extension.
        ''' The user is prompted to choose a file to import using an OpenFileDialog.
        ''' The settings are read from the file and sent to Grbl using the G-code syntax "ID=Value".
        ''' The progress of the import is displayed on a ProgressBar control.
        ''' </summary>
        ''' <remarks>
        ''' This function updates the _paramTable DataTable with the newly imported settings.
        ''' </remarks>
        ''' <exception cref="System.Exception">Thrown if there is an error during the import process.</exception>
        ''' <param name="progressBar">A ProgressBar control to display the import progress. This can be null if not used.</param>
        Public Sub ImportSettingsFromFile(progressBar As ProgressBar)
            ' Check if dgGrblSettings has rows
            If _gui.dgGrblSettings.Rows.Count = 0 Then
                MessageBox.Show("Grbl may not be connected! Settings should be initialized first")
                Return
            End If
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Grbl settings files (*.grbls)|*.grbls"

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                progressBar.Visible = True
                Try
                    Dim filePath As String = openFileDialog.FileName
                    ' Check if the file has a valid format
                    If Not IsValidGrblSettingsFile(filePath) Then
                        MessageBox.Show("Invalid file format. The file must contain settings in the format: $ID=Value")
                        Return
                    End If
                    ' Create a new DataTable to store the imported settings temporarily
                    Dim importedParamTable As New DataTable
                    With importedParamTable
                        .Columns.Add("ID")
                        .Columns.Add(Resources.GrblSettings_FillSettings_Value)
                        .PrimaryKey = New DataColumn() { .Columns("ID")}
                    End With

                    ' Get the total number of lines in the file to calculate the progress bar step
                    Dim totalLines As Integer = File.ReadAllLines(filePath).Length
                    Dim stepValue As Integer = If(totalLines > 0, 100 \ totalLines, 1)
                    Dim lineNumber As Integer = 0

                    ' Create a StreamReader to read the settings data from the file
                    Using reader As New StreamReader(filePath)
                        ' Read each line from the file and add it as a new row to the importedParamTable
                        While Not reader.EndOfStream
                            Dim paramLine As String = reader.ReadLine()
                            Dim paramParts As String() = paramLine.Split("="c)
                            If paramParts.Length = 2 Then
                                Dim paramID As String = paramParts(0)
                                Dim paramValue As String = paramParts(1)
                                importedParamTable.Rows.Add(paramID, paramValue)
                            End If
                        End While
                    End Using

                    ' Iterate through each setting and send it to Grbl
                    For Each row As DataRow In importedParamTable.Rows
                        Dim param As String = row("ID").ToString() & "=" & row(Resources.GrblSettings_FillSettings_Value).ToString()

                        ' Send the setting to Grbl
                        gcode.sendGCodeLine(param)

                        ' Wait for a short delay (adjust the delay time as needed)
                        Threading.Thread.Sleep(200) ' 200 milliseconds - adjust this value if needed

                        ' Update the progress bar
                        lineNumber += 1
                        progressBar.Value = Math.Min(lineNumber * stepValue, 100)
                    Next

                    ' Refresh the GrblSettings._paramTable with the newly imported settings
                    _paramTable.Rows.Clear()
                    For Each row As DataRow In importedParamTable.Rows
                        _paramTable.Rows.Add(row("ID"), row(Resources.GrblSettings_FillSettings_Value), "")
                    Next

                    ' Update the DataGridView with the newly imported settings
                    _gui.dgGrblSettings.DataSource = _paramTable
                    _gui.dgGrblSettings.Refresh()

                    ' Send the "$$" command to Grbl to refresh the settings
                    gcode.sendGCodeLine("$$")

                    progressBar.Value = 100
                    MessageBox.Show("Successfully imported settings!")
                    ' Reset the progress bar
                    progressBar.Value = 0
                    progressBar.Visible = False
                Catch ex As Exception
                    ' Reset the progress bar
                    progressBar.Value = 0
                    progressBar.Visible = False
                    MessageBox.Show("Failed to import settings from file : " & Environment.NewLine & ex.Message)
                End Try
            End If
        End Sub

    End Class
    ''' <summary>
    ''' Checks if the Grbl settings file has a valid format before importing.
    ''' The file is considered valid if it contains settings in the format: $ID=Value
    ''' </summary>
    ''' <param name="filePath">The file path of the Grbl settings file to be checked.</param>
    ''' <returns>True if the file has a valid format, False otherwise.</returns>
    Private Shared Function IsValidGrblSettingsFile(filePath As String) As Boolean
        Try
            ' Read all lines from the file
            Dim fileLines As String() = File.ReadAllLines(filePath)

            ' Check if the file format is valid
            For Each line As String In fileLines
                If Not line.StartsWith("$") Then
                    Return False
                End If
                Dim parts As String() = line.Split("="c)
                If parts.Length <> 2 Then
                    Return False
                End If
                Dim paramID As Integer
                If Not Integer.TryParse(parts(0).Substring(1), paramID) Then
                    Return False
                End If
            Next

            ' If we reach this point, the file format is valid
            Return True

        Catch ex As Exception
            ' An exception occurred during file reading or parsing
            ' The file is considered invalid in this case
            Return False
        End Try
    End Function

    ' We need a way to propogate changes on the Settings tab, do that here

    Private Sub btnSettingsRefreshMisc_Click(sender As Object, e As EventArgs) Handles btnSettingsRefreshMisc.Click, btnSettingsRefreshPosition.Click, btnSettingsRefreshJogging.Click
        Dim b As Button = sender
        Select Case DirectCast(b.Tag, String)
            Case "Misc"
                changeStatusRate(My.Settings.StatusPollInterval)
                prgBarQ.Maximum = My.Settings.QBuffMaxSize
                prgbRxBuf.Maximum = My.Settings.RBuffMaxSize
                cbStatusPollEnable.Checked = My.Settings.StatusPollEnabled
                cbSettingsConnectOnLoad.Checked = My.Settings.GrblConnectOnLoad

            Case "Position"
                tbSettingsSpclPosition1.Text = My.Settings.MachineSpclPosition1
                tbSettingsSpclPosition2.Text = My.Settings.MachineSpclPosition2


            Case "Jogging"
                tbSettingsFIImperial.Text = My.Settings.JoggingFIImperial
                tbSettingsFRImperial.Text = My.Settings.JoggingFRImperial
                tbSettingsFIMetric.Text = My.Settings.JoggingFIMEtric
                tbSettingsFRMetric.Text = My.Settings.JoggingFRMetric
                cbSettingsMetric.Checked = My.Settings.JoggingUnitsMetric

                setDistanceMetric(cbSettingsMetric.Checked)

                btnXPlus.Interval = 1000 / My.Settings.JoggingXRepeat
                btnXMinus.Interval = 1000 / My.Settings.JoggingXRepeat
                btnYPlus.Interval = 1000 / My.Settings.JoggingYRepeat
                btnYMinus.Interval = 1000 / My.Settings.JoggingYRepeat
                btnZPlus.Interval = 1000 / My.Settings.JoggingZRepeat
                btnZMinus.Interval = 1000 / My.Settings.JoggingZRepeat
        End Select
    End Sub


    Private Sub btnSettingsGrbl_Click(sender As Object, e As EventArgs) Handles btnSettingsGrbl.Click
        ' Retrieve settings from Grbl
        settings.RefreshSettings()
    End Sub

    Private Sub dgGrblSettings_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgGrblSettings.CellMouseDoubleClick
        ' User wants to edit a Grbl setting
        ' We do these one at a time due to EEProm write restrictions
        'Psuedo code:
        ' Which row? Get new value, determine $id, send to Grbl, do a refresh ($$)
        If e.ColumnIndex <> 1 Then
            ' ignore the click, it is not in Value column
            Return
        End If
        Dim gridView As DataGridView = sender
        If Not IsNothing(gridView.EditingControl) Then
            ' we have something to change (aka ignore errant double clicks)
            Dim param As String = gridView.Rows(e.RowIndex).Cells(0).Value.ToString & "=" & gridView.EditingControl.Text
            gcode.sendGCodeLine(param)
            Sleep(200)              ' Have to wait for EEProm write
            gcode.sendGCodeLine("$$")   ' Refresh
        End If
    End Sub

End Class
