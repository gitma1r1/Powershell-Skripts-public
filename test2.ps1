Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data

# CsvFile-Pfad
$CsvFilePath = "C:\Users\mai156\Desktop\audiobook_test_02\test.csv"

# Erstellen des Hauptfensters
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(900, 600)
$Form.StartPosition = "CenterScreen"
$Form.Text = "Tabelle"

# DataGridView erstellen
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Size = New-Object System.Drawing.Size(880, 500)
$DataGridView.AllowUserToAddRows = $false
$DataGridView.MultiSelect = $false

# Variable zur Verfolgung von Änderungen
$ChangesMade = $false

# Funktion zum Exportieren der Tabelle in eine CSV-Datei
function Export-DataGridViewToCSV {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView,
        [string]$FilePath
    )

    $DataTable = $DataGridView.DataSource

    # DataTable in CSV exportieren
    $DataTable | Export-Csv -Path $FilePath -NoTypeInformation -Append
}

# Funktion zum Importieren der Tabelle aus einer CSV-Datei
function Import-DataGridViewFromCSV {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView,
        [string]$FilePath
    )

    $CsvData = Import-Csv -Path $FilePath

    $DataTable = New-Object System.Data.DataTable

    $null = $DataTable.Columns.Add("Nr", [int])
    $null = $DataTable.Columns.Add("Titel")
    $null = $DataTable.Columns.Add("Author")
    $null = $DataTable.Columns.Add("Verlag")
    $null = $DataTable.Columns.Add("Sprecher")
    $null = $DataTable.Columns.Add("Dauer")
    $null = $DataTable.Columns.Add("Notizen")

    $rowIndex = 1
    foreach ($item in $CsvData) {
        $dataRow = $DataTable.NewRow()
        $dataRow["Nr"] = $rowIndex++
        $dataRow["Titel"] = $item."Titel"
        $dataRow["Author"] = $item."Author"
        $dataRow["Verlag"] = $item."Verlag"
        $dataRow["Sprecher"] = $item."Sprecher"
        $dataRow["Dauer"] = $item."Dauer"
        $dataRow["Notizen"] = $item."Notizen"
        $null = $DataTable.Rows.Add($dataRow)
    }

    # Hinzufügen der Daten zur Tabelle
    $DataGridView.DataSource = $DataTable
}

# Funktion zum Laden einer neuen Tabelle
function LoadTable {
    if ($ChangesMade) {
        $DialogResult = [System.Windows.Forms.MessageBox]::Show("Möchten Sie die bestehende Tabelle speichern, bevor Sie eine neue Tabelle laden?", "Bestehende Tabelle speichern?", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($DialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            SaveTable
        }

        if ($DialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
            return
        }
    }

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Dateien (*.csv)|*.csv"
    $OpenFileDialog.Title = "Tabelle laden"
    if ($OpenFileDialog.ShowDialog() -eq 'OK') {
        $OpenFilePath = $OpenFileDialog.FileName

        # Löschen der vorhandenen Zellen und der Tabelle
        $DataGridView.DataSource = $null

        Import-DataGridViewFromCSV -DataGridView $DataGridView -FilePath $OpenFilePath
        [System.Windows.Forms.MessageBox]::Show("Die Tabelle wurde erfolgreich geladen.", "Laden abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $ChangesMade = $false
        $Form.Text = "Tabelle"
        $SaveButton.Enabled = $false
    }
}

# Funktion zum Speichern der aktuellen Tabelle
function SaveTable {
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "CSV Dateien (*.csv)|*.csv"
    $SaveFileDialog.Title = "Tabelle speichern"
    if ($SaveFileDialog.ShowDialog() -eq 'OK') {
        $SaveFilePath = $SaveFileDialog.FileName

        if ([System.IO.Path]::GetExtension($SaveFilePath) -ne ".csv") {
            [System.Windows.Forms.MessageBox]::Show("Ungültige Dateiendung. Bitte wählen Sie eine CSV-Datei aus.", "Fehler beim Speichern", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Aktualisieren des DataTables mit den Änderungen aus dem DataGridView
        $DataTable = $DataGridView.DataSource
        $DataTable.AcceptChanges()

        # Exportieren der aktualisierten Tabelle in eine CSV-Datei
        $DataTable | ConvertTo-Csv -NoTypeInformation | Set-Content -Path $SaveFilePath
        [System.Windows.Forms.MessageBox]::Show("Die Tabelle wurde erfolgreich gespeichert.", "Speichern abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $ChangesMade = $false
        $Form.Text = "Tabelle"
        $SaveButton.Enabled = $false
    }
}

# Funktion zum Hinzufügen einer neuen Zeile zur Tabelle
function AddRow {
    $DataTable = $DataGridView.DataSource
    $newRow = $DataTable.NewRow()
    $newRow["Nr"] = $DataTable.Rows.Count + 1
    $newRow["Titel"] = ""
    $newRow["Author"] = ""
    $newRow["Verlag"] = ""
    $newRow["Sprecher"] = ""
    $newRow["Dauer"] = ""
    $newRow["Notizen"] = ""
    $DataTable.Rows.Add($newRow)
    $DataGridView.CurrentCell = $DataGridView.Rows[$DataTable.Rows.Count - 1].Cells[1]
    $DataGridView.BeginEdit($true)
    $ChangesMade = $true
    $Form.Text = "Tabelle*"
    $SaveButton.Enabled = $true
}

# Funktion zum Löschen der ausgewählten Zeilen aus der Tabelle
function DeleteSelectedRows {
    $DataTable = $DataGridView.DataSource
    $selectedRows = $DataGridView.SelectedRows

    if ($selectedRows.Count -gt 0) {
        $result = [System.Windows.Forms.MessageBox]::Show("Möchten Sie die ausgewählten Zeilen wirklich löschen?", "Zeilen löschen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            foreach ($row in $selectedRows) {
                $DataTable.Rows.RemoveAt($row.Index)
            }

            # Aktualisieren der Nr-Spalte
            $i = 1
            foreach ($row in $DataTable.Rows) {
                $row["Nr"] = $i++
            }

            $ChangesMade = $true
            $Form.Text = "Tabelle*"
            $SaveButton.Enabled = $true
        }
    }

    $DataGridView.ClearSelection()
}

# Funktion zum Überprüfen, ob Änderungen an der Tabelle vorgenommen wurden
function HasChanges {
    $DataTable = $DataGridView.DataSource
    return $DataTable.GetChanges() -ne $null
}

# DataGridView CellValueChanged Event Handler
$CellValueChanged = {
    if (!$ChangesMade) {
        $ChangesMade = $true
        $Form.Text = "Tabelle*"
        $SaveButton.Enabled = $true
    }
}

# DataGridView CurrentCellDirtyStateChanged Event Handler
$CurrentCellDirtyStateChanged = {
    if ($DataGridView.IsCurrentCellDirty) {
        $DataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
}

# DataGridView SelectionChanged Event Handler
$SelectionChanged = {
    if ($DataGridView.SelectedCells.Count -gt 0) {
        $DataGridView.BeginEdit($false)
    }
}

# DataGridView CellBeginEdit Event Handler
$CellBeginEdit = {
    $Form.Text = "Tabelle*"
}

# DataGridView CellEndEdit Event Handler
$CellEndEdit = {
    if (HasChanges) {
        $Form.Text = "Tabelle*"
    } else {
        $Form.Text = "Tabelle"
    }
}

# DataGridView KeyDown Event Handler
$KeyDown = {
    param($sender, $e)

    if ($e.Control -and $e.KeyCode -eq "S") {
        $e.SuppressKeyPress = $true
        SaveTable
    }
}

# DataGridView KeyPress Event Handler
$KeyPress = {
    param($sender, $e)

    if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Control -and $e.KeyChar -eq [System.Windows.Forms.Keys]::S) {
        $e.Handled = $true
    }
}

# DataGridView CellValidating Event Handler
$CellValidating = {
    param($sender, $e)

    if ($e.ColumnIndex -eq 0 -and $e.FormattedValue -eq "") {
        $e.Cancel = $true
        $DataGridView.Rows[$e.RowIndex].ErrorText = "Die Nummer darf nicht leer sein."
    }
}

# DataGridView CellValidated Event Handler
$CellValidated = {
    param($sender, $e)

    if ($e.ColumnIndex -eq 0) {
        $DataGridView.Rows[$e.RowIndex].ErrorText = ""
    }
}

# DataGridView CellMouseClick Event Handler
$CellMouseClick = {
    param($sender, $e)

    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $DataGridView.ClearSelection()
        $DataGridView.Rows[$e.RowIndex].Selected = $true
        $DataGridView.CurrentCell = $DataGridView.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
    }
}

# DataGridView CellParsing Event Handler
$CellParsing = {
    param($sender, $e)

    if ($e.ColumnIndex -eq 0) {
        try {
            $value = [int]$e.Value
            $e.ParsingApplied = $true
        } catch {
            $e.ParsingApplied = $false
            [System.Windows.Forms.MessageBox]::Show("Die eingegebene Nummer hat das falsche Format. Bitte geben Sie eine ganze Zahl ein.", "Fehler beim Parsen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
}

# DataGridView ContextMenuStrip
$ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip

# ContextMenuStrip - Löschen
$DeleteMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$DeleteMenuItem.Text = "Zeile löschen"
$DeleteMenuItem.Add_Click({ DeleteSelectedRows })
$ContextMenuStrip.Items.Add($DeleteMenuItem)

# DataGridView CellContextMenuStripNeeded Event Handler
$CellContextMenuStripNeeded = {
    param($sender, $e)

    if ($e.RowIndex -ge 0) {
        $DataGridView.CurrentCell = $DataGridView.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
        $DataGridView.Rows[$e.RowIndex].Selected = $true
        $DataGridView.ContextMenuStrip = $ContextMenuStrip
    } else {
        $DataGridView.ContextMenuStrip = $null
    }
}

# DataGridView Events hinzufügen
$DataGridView.Add_CellValueChanged($CellValueChanged)
$DataGridView.Add_CurrentCellDirtyStateChanged($CurrentCellDirtyStateChanged)
$DataGridView.Add_SelectionChanged($SelectionChanged)
$DataGridView.Add_CellBeginEdit($CellBeginEdit)
$DataGridView.Add_CellEndEdit($CellEndEdit)
$DataGridView.Add_KeyDown($KeyDown)
$DataGridView.Add_KeyPress($KeyPress)
$DataGridView.Add_CellValidating($CellValidating)
$DataGridView.Add_CellValidated($CellValidated)
$DataGridView.Add_CellMouseClick($CellMouseClick)
$DataGridView.Add_CellParsing($CellParsing)
$DataGridView.Add_CellContextMenuStripNeeded($CellContextMenuStripNeeded)

# Button zum Speichern der Tabelle
$SaveButton = New-Object System.Windows.Forms.Button
$SaveButton.Location = New-Object System.Drawing.Point(10, 520)
$SaveButton.Size = New-Object System.Drawing.Size(100, 30)
$SaveButton.Text = "Tabelle speichern"
$SaveButton.Enabled = $false
$SaveButton.Add_Click({ SaveTable })
$Form.Controls.Add($SaveButton)

# Button zum Laden der Tabelle
$LoadButton = New-Object System.Windows.Forms.Button
$LoadButton.Location = New-Object System.Drawing.Point(120, 520)
$LoadButton.Size = New-Object System.Drawing.Size(100, 30)
$LoadButton.Text = "Tabelle laden"
$LoadButton.Add_Click({ LoadTable })
$Form.Controls.Add($LoadButton)

# Button zum Hinzufügen einer neuen Zeile
$AddRowButton = New-Object System.Windows.Forms.Button
$AddRowButton.Location = New-Object System.Drawing.Point(230, 520)
$AddRowButton.Size = New-Object System.Drawing.Size(100, 30)
$AddRowButton.Text = "Zeile hinzufügen"
$AddRowButton.Add_Click({ AddRow })
$Form.Controls.Add($AddRowButton)

# Button zum Löschen der ausgewählten Zeilen
$DeleteRowsButton = New-Object System.Windows.Forms.Button
$DeleteRowsButton.Location = New-Object System.Drawing.Point(340, 520)
$DeleteRowsButton.Size = New-Object System.Drawing.Size(100, 30)
$DeleteRowsButton.Text = "Zeile löschen"
$DeleteRowsButton.Enabled = $true
$DeleteRowsButton.Add_Click({ DeleteSelectedRows })
$Form.Controls.Add($DeleteRowsButton)

# Button zum Zurückscrollen zur ersten Zeile
$ScrollToTopButton = New-Object System.Windows.Forms.Button
$ScrollToTopButton.Location = New-Object System.Drawing.Point(450, 520)
$ScrollToTopButton.Size = New-Object System.Drawing.Size(100, 30)
$ScrollToTopButton.Text = "Zum Anfang"
$ScrollToTopButton.Enabled = $false
$ScrollToTopButton.Add_Click({
    $DataGridView.FirstDisplayedScrollingRowIndex = 0
})
$Form.Controls.Add($ScrollToTopButton)

# DataGridView Scroll Event Handler
$Scroll = {
    if ($DataGridView.FirstDisplayedScrollingRowIndex -eq 0) {
        $ScrollToTopButton.Enabled = $false
    } else {
        $ScrollToTopButton.Enabled = $true
    }
}
$DataGridView.Add_Scroll($Scroll)


# Import CSV from CsvFilePath
Import-DataGridViewFromCSV -DataGridView $DataGridView -FilePath $CsvFilePath

# Hinzufügen der Tabelle zum Hauptfenster
$Form.Controls.Add($DataGridView)

# Anzeigen des Hauptfensters
$Form.Add_FormClosing({
    if ($ChangesMade) {
        $result = [System.Windows.Forms.MessageBox]::Show("Möchten Sie die Änderungen an der Tabelle speichern, bevor Sie das Programm beenden?", "Änderungen speichern?", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            SaveTable
        }

        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            $Form.Cancel = $true
        }
    }
})

$Form.ShowDialog() | Out-Null
