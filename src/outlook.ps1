

$OL=New-Object -ComObject OUTLOOK.APPLICATION
$ns = $OL.GETNAMESPACE("MAPI")


# Last email
#$mail = $ns.Folders.Item(1).Folders.Item("Posteingang").Items(1)
# | Select-Object -Property Subject : Get who
# .Attachments


#$OL.AdvancedSearch("Posteingang","urn:schemas:httpmail:hasattachment=true ").Results[1]
#$mail = $ns.Folders.Item(1).Folders.Item("Posteingang").Items

$date = Get-Date (Get-Date).AddDays(-2) -Format 'dd/MM/yyyy'
$filter = "[ReceivedTime] >= '$date 18:00'"
$allmails = $ns.Folders.Item(1).Folders.Item("Posteingang").Items.Restrict($filter)


# Attachments of last mail
#$mail[1].Attachments

# Name 
#$mail[1].Attachments[1].FileName

# For one mail, get all true attachments
#$mail[3] | Where-Object { $_.Attachments.FileName -notmatch "(image[0-9][0-9][0-9]|.msg)" }
#$goodmails = ( $mail[3] | Where-Object { $_.Attachments.FileName -match "(.pdf|.doc|.xls)" } )




foreach ($mail in $allmails)
{

    #[bool]$istrueattach = $false
    foreach ( $attach in $mail.Attachments ) 
    {
        #echo $attach.FileName
        if ($attach.FileName -match  "(.pdf|.doc|.xls)" )
        {
            echo "found !"
            echo $mail.Subject
            echo $mail.SenderName

        }


    }

}




#echo $goodmails.Subject


