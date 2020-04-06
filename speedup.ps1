param ($SourceFile)
$f = Get-Item $SourceFile

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Computer'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Selecione a velocidade:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

[void] $listBox.Items.Add('0.1')
[void] $listBox.Items.Add('0.1666')
[void] $listBox.Items.Add('0.2')
[void] $listBox.Items.Add('0.5')
[void] $listBox.Items.Add('2')
[void] $listBox.Items.Add('4')
[void] $listBox.Items.Add('8')
[void] $listBox.Items.Add('10')
[void] $listBox.Items.Add('16')

$form.Controls.Add($listBox)

$form.Topmost = $true

$listBox.SelectedIndex = 7

$result = $form.ShowDialog()

$speed = 10


if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $speed = $listBox.SelectedItem
} else {
    return
}


$dstFile = $f.FullName.Replace($f.Extension, ("_speed_" + $speed + $f.Extension))

Write-Progress -Activity "MEncoder" -PercentComplete 0

C:\Tools\tencoder_64\MEncoder\MEncoder64\mencoder.exe -msglevel all=5:mencoder=0 -speed $speed -ofps 24 -nosound -o $dstFile -ovc lavc  $SourceFile | % {

  $ol = ("" + $_)

  if ($ol.StartsWith("Pos:"))
  {
    $ol | Select-String -Pattern "\(\d?\d?\d%\)" -AllMatches | % {
        $pct = $_.matches.value.substring(1, 2)
        Write-Progress -Activity "MEncoder" -PercentComplete $pct -CurrentOperation $dstFile
    }
  }
}

Write-Progress -Activity "MEncoder" -Completed
