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

# Button zum Speichern der Tabelle
$SaveButton = New-Object System.Windows.Forms.Button
$SaveButton.Location = New-Object System.Drawing.Point(10, 520)
$SaveButton.Size = New-Object System.Drawing.Size(100, 30)
$SaveButton.Text = "Tabelle speichern"
$SaveButton.Add_Click({
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "CSV Dateien (*.csv)|*.csv"
    $SaveFileDialog.Title = "Tabelle speichern"
    $SaveFileDialog.ShowDialog() | Out-Null
    $SaveFilePath = $SaveFileDialog.FileName

    if ($SaveFilePath) {
        Export-DataGridViewToCSV -DataGridView $DataGridView -FilePath $SaveFilePath
        [System.Windows.Forms.MessageBox]::Show("Die Tabelle wurde erfolgreich gespeichert.", "Speichern abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$Form.Controls.Add($SaveButton)

# Button zum Laden der Tabelle
$LoadButton = New-Object System.Windows.Forms.Button
$LoadButton.Location = New-Object System.Drawing.Point(120, 520)
$LoadButton.Size = New-Object System.Drawing.Size(100, 30)
$LoadButton.Text = "Tabelle laden"
$LoadButton.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Dateien (*.csv)|*.csv"
    $OpenFileDialog.Title = "Tabelle laden"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFilePath = $OpenFileDialog.FileName

    if ($OpenFilePath) {
        # Löschen der vorhandenen Zellen und der Tabelle
        $DataGridView.DataSource = $null

        Import-DataGridViewFromCSV -DataGridView $DataGridView -FilePath $OpenFilePath

        [System.Windows.Forms.MessageBox]::Show("Die Tabelle wurde erfolgreich geladen.", "Laden abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$Form.Controls.Add($LoadButton)

# Funktion zum Exportieren der Tabelle in eine CSV-Datei
function Export-DataGridViewToCSV {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView,
        [string]$FilePath
    )

    $DataTable = New-Object System.Data.DataTable

    # Spalten hinzufügen
    foreach ($column in $DataGridView.Columns) {
        $null = $DataTable.Columns.Add($column.Name)
    }

    # Zeilen hinzufügen
    foreach ($row in $DataGridView.Rows) {
        $dataRow = $DataTable.NewRow()
        for ($i = 0; $i -lt $DataGridView.ColumnCount; $i++) {
            $dataRow[$i] = $row.Cells[$i].Value
        }
        $null = $DataTable.Rows.Add($dataRow)
    }

    # DataTable in CSV exportieren
    $DataTable | Export-Csv -Path $FilePath -NoTypeInformation
}

# Funktion zum Importieren der Tabelle aus einer CSV-Datei
function Import-DataGridViewFromCSV {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView,
        [string]$FilePath
    )

    $DataTable = New-Object System.Data.DataTable
    $null = $DataTable.Columns.Add("Nr")
    $null = $DataTable.Columns.Add("Titel")
    $null = $DataTable.Columns.Add("Author")
    $null = $DataTable.Columns.Add("Verlag")
    $null = $DataTable.Columns.Add("Sprecher")
    $null = $DataTable.Columns.Add("Dauer")
    $null = $DataTable.Columns.Add("Notizen")

    $CsvData = Import-Csv -Path $FilePath

    foreach ($CsvRow in $CsvData) {
        $DataRow = $DataTable.NewRow()
        $DataRow["Nr"] = $CsvRow."Nr"
        $DataRow["Titel"] = $CsvRow."Titel"
        $DataRow["Author"] = $CsvRow."Author"
        $DataRow["Verlag"] = $CsvRow."Verlag"
        $DataRow["Sprecher"] = $CsvRow."Sprecher"
        $DataRow["Dauer"] = $CsvRow."Dauer"
        $DataRow["Notizen"] = $CsvRow."Notizen"
        $DataTable.Rows.Add($DataRow)
    }

    # Hinzufügen der Daten zur Tabelle
    $DataGridView.DataSource = $DataTable
}

# Import CSV from CsvFilePath
Import-DataGridViewFromCSV -DataGridView $DataGridView -FilePath $CsvFilePath

# Hinzufügen der Tabelle zum Hauptfenster
$Form.Controls.Add($DataGridView)

# Anzeigen des Hauptfensters
$Form.ShowDialog() | Out-Null
