<#
.SYNOPSIS
    Demonstration of System.Windows.Forms.ListView sorting via PowerShell.
.DESCRIPTION
    Displays a basic Windows Form with a ListView control and some items. Each column is sortable.
.LINK
    https://etechgoodness.wordpress.com/2014/02/25/sort-a-windows-forms-listview-in-powershell-without-a-custom-comparer/
.NOTES
    By Eric Siron
    Version 1.0.1 June 6th 2014:  Slightly modified to work with a wider range of PS hosts and versions.
    Version 1.1 August 12th 2015: Improved test procedure per feedback from reader Stanley
#>
 
## Set up the environment
Add-Type -AssemblyName System.Windows.Forms
$LastColumnClicked = 0 # tracks the last column number that was clicked
$LastColumnAscending = $false # tracks the direction of the last sort of this column
 
## Create a form and a ListView
$Form = New-Object System.Windows.Forms.Form
$ListView = New-Object System.Windows.Forms.ListView
 
## Configure the form
$Form.Text = "ListView Sort Demo"
 
## Configure the ListView
$ListView.View = [System.Windows.Forms.View]::Details
$ListView.Width = $Form.ClientRectangle.Width
$ListView.Height = $Form.ClientRectangle.Height
$ListView.Anchor = "Top, Left, Right, Bottom"
 
# Add the ListView to the Form
$Form.Controls.Add($ListView)
 
# Add columns to the ListView
$ListView.Columns.Add("Item Name", -2) | Out-Null
$ListView.Columns.Add("Color") | Out-Null
$ListView.Columns.Add("Size") | Out-Null
$ListView.Columns.Add("Weight") | Out-Null
 
# Add list items
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Barlav")
$ListViewItem.Subitems.Add("Green") | Out-Null
$ListViewItem.Subitems.Add("Tiny") | Out-Null
$ListViewItem.Subitems.Add("11") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Floomquet")
$ListViewItem.Subitems.Add("Red") | Out-Null
$ListViewItem.Subitems.Add("Large") | Out-Null
$ListViewItem.Subitems.Add("2") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Gardgel")
$ListViewItem.Subitems.Add("Yellow") | Out-Null
$ListViewItem.Subitems.Add("Jumbo") | Out-Null
$ListViewItem.Subitems.Add("7") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Wilbit")
$ListViewItem.Subitems.Add("Turquoise") | Out-Null
$ListViewItem.Subitems.Add("Gigantic") | Out-Null
$ListViewItem.Subitems.Add("1") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Zasitch")
$ListViewItem.Subitems.Add("Beige") | Out-Null
$ListViewItem.Subitems.Add("Small") | Out-Null
$ListViewItem.Subitems.Add("5") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
## Set up the event handler
$ListView.add_ColumnClick({SortListView $_.Column})
 
## Event handler
function SortListView
{
 param([parameter(Position=0)][UInt32]$Column)
 
$Numeric = $true # determine how to sort
 
# if the user clicked the same column that was clicked last time, reverse its sort order. otherwise, reset for normal ascending sort
if($Script:LastColumnClicked -eq $Column)
{
    $Script:LastColumnAscending = -not $Script:LastColumnAscending
}
else
{
    $Script:LastColumnAscending = $true
}
$Script:LastColumnClicked = $Column
$ListItems = @(@(@())) # three-dimensional array; column 1 indexes the other columns, column 2 is the value to be sorted on, and column 3 is the System.Windows.Forms.ListViewItem object
 
foreach($ListItem in $ListView.Items)
{
    # if all items are numeric, can use a numeric sort
    if($Numeric -ne $false) # nothing can set this back to true, so don't process unnecessarily
    {
        try
        {
            $Test = [Double]$ListItem.SubItems[[int]$Column].Text
        }
        catch
        {
            $Numeric = $false # a non-numeric item was found, so sort will occur as a string
        }
    }
    $ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)
}
 
# create the expression that will be evaluated for sorting
$EvalExpression = {
    if($Numeric)
    { return [Double]$_[0] }
    else
    { return [String]$_[0] }
}
 
# all information is gathered; perform the sort
$ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=$Script:LastColumnAscending}
 
## the list is sorted; display it in the listview
$ListView.BeginUpdate()
$ListView.Items.Clear()
foreach($ListItem in $ListItems)
{
    $ListView.Items.Add($ListItem[1])
}
$ListView.EndUpdate()
}
 
## Show the form
$Response = $Form.ShowDialog()