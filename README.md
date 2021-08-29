# tridate
Sub TriDates()
Set NS = olapp.GetNamespace("MAPI")
Set dossier = NS.Folders("meteocaudry@gmail.com").Folders("Boîte de réception")

l = 0
k = 0

If Not Len(Dir("C:\Users\mikl\Desktop\meteo\", vbDirectory)) > 0 Then
        MkDir "C:\Users\mikl\Desktop\meteo\"
    End If
    

dossiersave = "C:\Users\mikl\Desktop\meteo\"

For Each i In dossier.Items
    ReceivedTime = Left(i.ReceivedTime, 10) ' isole la date
    ReceivedTimecorrige = Replace(ReceivedTime, "/", "")
    
    dossiersave = "C:\Users\mikl\Desktop\meteo\" + ReceivedTimecorrige + "\"
    
    If Not Len(Dir(dossiersave, vbDirectory)) > 0 Then
        MkDir dossiersave
    End If
        
    Set objAttachments = i.Attachments
    lngCount = objAttachments.Count
    ' mettre la piece jointe dans le dossier
    For k = lngCount To 1 Step -1
    strFile = objAttachments.Item(k).Filename

    ' Combine with the path to the Temp folder.
     strFile = dossiersave & strFile

    ' Save the attachment as a file.
     objAttachments.Item(k).SaveAsFile strFile
    
    Next k
Next i
End Sub
