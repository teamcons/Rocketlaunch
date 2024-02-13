

$OL=New-Object -ComObject OUTLOOK.APPLICATION
$ns =$OL.GETNAMESPACE("MAPI")


# Last email
$mail = $ns.Folders.Item(1).Folders.Item("Posteingang").Items(1)
# | Select-Object -Property Subject : Get who
# .Attachments


#$OL.AdvancedSearch("Posteingang","urn:schemas:httpmail:hasattachment=true ").Results[1]
$mail = $ns.Folders.Item(1).Folders.Item("Posteingang").Items

$date = Get-Date (Get-Date).AddDays(-7) -Format 'd/MM/yyyy hh.mm tt'
$filter = "[ReceivedTime] >= '$date'"
$mail = $ns.Folders.Item(1).Folders.Item("Posteingang").Items.Restrict($filter)


# Attachments of last mail
$mail[1].Attachments

# Name 
$mail[1].Attachments[1].FileName



.Items.Restrict('[UnRead] =   True')
$mail | Select-Object -Property Body | Format-List