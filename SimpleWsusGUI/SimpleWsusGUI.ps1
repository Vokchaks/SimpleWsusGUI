<#
    .SYNOPSIS
    SimpleWsusGUI it's not a substitute but a good helper WSUS service

    .DESCRIPTION
     For basis get command management  WSUS via powershell
     https://learn-powershell.net/2010/10/25/wsus-managing-groups-with-powershell/

    .EXAMPLE
    Run in Administration mode
    .\SimpleWsusGUI.ps1 

    .DATE
    03.04.2017
#>

function BranchList ( $branch, $group) {
    if (!$group) {
        return
    }

    Foreach ($local:g in $group.GetChildTargetGroups()) {
        $local:b = $branch.Nodes.Add($local:g.Name)
        BranchList  $local:b $local:g
    }
}



$monitor = [System.Windows.Forms.Screen]::PrimaryScreen

$ScreenWidth = $monitor.WorkingArea.Width
$ScreenHeight = $monitor.WorkingArea.Height

Add-Type -assembly System.Windows.Forms
$MainForm = New-Object System.Windows.Forms.Form

$dataGridView = $null

$MainForm.Text = 'GUI WSUS Console'

$MainForm.Width = 800
$MainForm.Height = 600
$MainForm.AutoSize = $true
$MainForm.StartPosition = "CenterScreen"

if($ScreenWidth -lt 800 -or $ScreenHeight -lt 600){
    $MainForm.WindowState = 'Maximized'
}

$StatusBar = New-Object System.Windows.Forms.StatusBar
$StatusBarPanel = New-Object System.Windows.Forms.StatusBarPanel
$StatusBarPanel.AutoSize = [System.Windows.Forms.StatusBarPanelAutoSize]::Contents
$StatusBarPanel.text = "Ready.."
$StatusBar.Panels.Add($StatusBarPanel)
$StatusBar.showpanels = $True 

$MainForm.Controls.Add($StatusBar)


$Label = New-Object System.Windows.Forms.Label
$Label.Text = "WSUS groups"
$Label.Location = New-Object System.Drawing.Point(10, 10)
$Label.AutoSize = $true
$MainForm.Controls.Add($Label)


$TreeView = New-Object System.Windows.Forms.TreeView
$TreeView.Location = New-Object System.Drawing.Point(10, 30)
$TreeView.Width = 250
$TreeView.Height = ($MainForm.Height - 40 - 60)
$TreeView.AutoSize = $true

$LabelComp = New-Object System.Windows.Forms.Label
$LabelComp.Text = "Computers: 0"
$LabelComp.Location = New-Object System.Drawing.Point(($TreeView.Width + 50), 10)
$LabelComp.AutoSize = $true
$MainForm.Controls.Add($LabelComp)

[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")            
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer("ngm-srv-008", $False, 8530)
$groups = $wsus.GetComputerTargetGroups()

$rebuidTree = $false

$allComps = ($groups | Where-Object {$_.Name -eq "Все компьютеры"}).GetComputerTargets()

$mainGroup = $null

Foreach ($group in $groups) {
    # $group.Name
    Try {
        $g = $group.GetParentTargetGroup()
    }
    Catch {
        $mainGroup = $group
    }
}

$TreeViewNode = $TreeView.Nodes.Add($mainGroup.Name)

BranchList  $TreeViewNode $mainGroup 

$MainForm.Controls.Add($TreeView)

$lastSelectedText = ""

$TreeView.Add_AfterSelect({
        $rebuildTree = $true
              
        $StatusBarPanel.text = $TreeView.SelectedNode.Text       


        if ($dataGridView -and ($lastSelectedText -ne $TreeView.SelectedNode.Text) ) {
            #$listView.Items.Clear()
            $dataGridView.Rows.Clear()

            foreach( $group in $groups) {
                if($group.Name -eq $TreeView.SelectedNode.Text ) {
                    $comps = $group.GetComputerTargets()
                    $LabelComp.Text = "Computers: " + $comps.Count
                    $row = 0
                    foreach ( $comp in $comps ) {
                        $dataGridView.Rows.Add($comp.FullDomainName, $comp.OSDescription, $comp.ClientVersion.ToString())
                        #$ids = $comp.ComputerTargetGroupIds()

                        foreach ($id in ($comp.ComputerTargetGroupIds).Guid) {                        
                            $dataGridView.Rows[$row].Cells[($groups |where-Object {$_.Id -eq $id}).Name].Value = $true
                        }
                        $row++

                        $dataGridView.AllowUserToAddRows = $false
                        $dataGridView.AllowUserToDeleteRows = $false
                        $dataGridView.Columns["Name"].ReadOnly = $true
                        $dataGridView.Columns["OS"].ReadOnly = $true
                        $dataGridView.Columns["Version"].ReadOnly = $true
                        $dataGridView.Columns["Все компьютеры"].ReadOnly = $true
                        $dataGridView.Columns["Неназначенные компьютеры"].ReadOnly = $true

                    }
                }
            }
        }
        $rebuildTree = $false
    })

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(($TreeView.Width + 50), 30)
$dataGridView.Size = New-Object System.Drawing.Size(($MainForm.Width - 60 - $TreeView.Width -40 ), ($MainForm.Height - 40 - 60))
#$dataGridView.AutoSize = $true
$MainForm.Controls.Add($dataGridView)

$dataGridView.ColumnCount = 3
$i = 0

$dataGridView.Columns[$i++].Name = 'Name'
$dataGridView.Columns[$i++].Name = 'Os'
$dataGridView.Columns[$i++].Name = 'Version'

Foreach ($group in $groups) {
    $StatusBarPanel.text = $i
    # Write-Host $i
    $dataGridView.Columns.Insert($i, (New-Object System.Windows.Forms.DataGridViewCheckBoxColumn))
    $dataGridView.Columns[$i++].Name = $group.Name
}


$dataGridView.Add_CurrentCellDirtyStateChanged({
    param($Sender,$EventArgs)

    if($Sender.IsCurrentCellDirty){
        $Sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

$dataGridView.Add_CellValueChanged({
    param($Sender,$EventArgs)

    if(!$rebuildTree) {
    
   
        $group = $groups | where-Object {$_.Name -eq $dataGridView.Columns[$EventArgs.ColumnIndex].Name}

         
        $comp = $allComps | Where-Object {$_.FullDomainName -eq $dataGridView.Rows[$EventArgs.RowIndex].Cells[0].Value}

        $inside = [bool](($groups | where-Object {$_.Name -eq $dataGridView.Columns[$EventArgs.ColumnIndex].Name}).GetComputerTargets() | Where-Object {$_.FullDomainName -eq $comp.FullDomainName})

        $StatusBarPanel.text = $comp.FullDomainName + " " + $inside

        # если истина (чекбокс отмечен) и комплютер не был в группе -> добавляем
        if ( $dataGridView.Rows[$EventArgs.RowIndex].Cells[$EventArgs.ColumnIndex].Value )
        { 
            if ( !$inside) {        
                $group.AddComputerTarget($comp)
                $StatusBarPanel.text = "Add " + $comp.FullDomainName + " to group: " + $group.Name
            }
        }
        # если ложь (чекбокс очищен) и компьютер был в группе - > удаляем
        else
        {
            if($inside)
            {
                $group.RemoveComputerTarget($comp)
                $StatusBarPanel.text = "Remove " + $comp.FullDomainName + " from group: " + $group.Name
            }
        }            
    }
})


$MainForm.Add_Resize({
    $TreeView.Height = ($MainForm.Height - 40 - 60 )    
    $dataGridView.Width = ($MainForm.Width - 60 - $TreeView.Width - 40 )
    $dataGridView.Height = ($MainForm.Height - 40 - 60)
})

$TreeView.ExpandAll()
$MainForm.ShowDialog()
