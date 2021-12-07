'Thunderbird contacts to outlook ( via LDIF )
'Crédit Olivier NOBLANC olivier.noblanc@dreets.gouv.Fr

Option Explicit

Dim LDIFArray(25000) As String

Sub LDIFF_REPARSE()
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False

Dim fd As Office.FileDialog
Set fd = xlApp.Application.FileDialog(msoFileDialogFilePicker)

Dim selectedItem As Variant
Dim a
If fd.Show = -1 Then
    For Each selectedItem In fd.SelectedItems
        Debug.Print selectedItem
        a = ReadtextFile(selectedItem)
    Next
End If

Set fd = Nothing
    xlApp.Quit
Set xlApp = Nothing


MsgBox (Decode_UTF8(Base64Decode("VGVybWluw6kgYnkgRVNJQyBCRkMgQ3LDqWRpdHMgOiBvbGl2aWVyLm5vYmxhbmNAZHJlZXRzLmdvdXYuZnI=")))

End Sub




Function ReadtextFile(FilePath) As Boolean
    Dim FileNum As Integer
    Dim DataLine As String
    
    FileNum = FreeFile()
    Open FilePath For Input As #FileNum
    Dim i As Integer
    i = 0
    While Not EOF(FileNum)
        Line Input #FileNum, DataLine ' read in data 1 line at a time
        Debug.Print (DataLine)
        If Len(DataLine) < 2 Then
            'éxécute la création
            executecrea
            Erase LDIFArray
            i = 0
        Else
           LDIFArray(i) = DataLine
           i = i + 1
        End If
        ' decide what to do with dataline,
        ' depending on what processing you need to do for each case
    Wend

    ReadtextFile = True
End Function

Private Sub executecrea()
  Dim detectionOK As Boolean
  detectionOK = False
  'je détecte si c'est une personne ou un groupe
  Dim item
   For Each item In LDIFArray
     If Len(item) < 2 Then
       Exit For
     End If
     If InStr(item, ":") Then
       If Trim(Split(item, ":")(1)) = "person" Then
          traitementcontactpersonne
          detectionOK = True
          Exit For
       End If
       If Trim(Split(item, ":")(1)) = "groupOfNames" Then
          traitementgroupes
          detectionOK = True
          Exit For
       End If
     End If
   Next
End Sub

Private Sub traitementgroupes()
    Dim nomgroupe As String
    Dim ligne As String
    
     Dim objOutlook As Outlook.Application
    'Dim myDistList As DistributionListItem
    Dim myDistList As Outlook.DistListItem
 
    'Crée l'instance Outlook
    Set objOutlook = New Outlook.Application
    'Crée un élément pour les contacts
    Set myDistList = objOutlook.CreateItem(olDistributionListItem)
    Dim myTempItem As Outlook.MailItem
    Dim myRecipients As Outlook.Recipients
    Set myTempItem = objOutlook.CreateItem(olMailItem)
    Set myRecipients = myTempItem.Recipients
    
    Dim addrmailcont As String
    Dim item
    For Each item In LDIFArray
        ligne = item
        ligne = traitementligne(ligne)
        If Len(item) < 2 Then
          Exit For
        End If
        If InStr(ligne, ":") Then
            If Trim(Split(ligne, ":")(0)) = "cn" Then
                     If isUTF8(ligne) Then
                        ligne = Decode_UTF8(ligne)
                     End If
                     myDistList.DLName = Trim(Split(ligne, ":")(1))
            End If
            If Trim(Split(ligne, ":")(0)) = "member" Then
                  addrmailcont = Trim(Split(ligne, "mail=")(1))
                  addrmailcont = Replace(addrmailcont, "'", "")
                  If InStr(addrmailcont, ",") Then
                    addrmailcont = Trim(Split(ligne, ",")(0))
                  End If
               'ajouter le contact à la liste
                myRecipients.Add addrmailcont
            End If
            
        End If
   Next
   
   myDistList.AddMembers myRecipients
   ' Sauvegarder la liste
   myDistList.Save
   
End Sub

Private Sub traitementcontactpersonne()
    Dim persFound As Boolean
    persFound = False
    Dim item
     For Each item In LDIFArray
     If InStr(item, ":") Then
       If Trim(Split(item, ":")(0)) = "mail" Then
          persFound = findcont(Trim(Split(item, ":")(1)))
          If persFound = False Then
            ajouterContactOutlook
          End If
          Exit For
       End If
     End If
   Next
End Sub

Private Sub ajouterContactOutlook()
    'Nécessite d'activer la référence https://excel.developpez.com/faq/?page=Messagerie#AjouterContact
        'Microsoft Outlook xx.x Object Library
    Dim objOutlook As Outlook.Application
    Dim objContact As ContactItem
 
    'Crée l'instance Outlook
    Set objOutlook = New Outlook.Application
    'Crée un élément pour les contacts
    Set objContact = objOutlook.CreateItem(olContactItem)
    Dim ligne As String
    Dim item
     For Each item In LDIFArray
        If Len(item) < 2 Then
            Exit For
        End If
        ligne = item
        ligne = traitementligne(ligne)
        
        If InStr(ligne, ":") Then
            'objContact.homest
            If isUTF8(ligne) Then
               ligne = Decode_UTF8(ligne)
            End If
            
            If Trim(Split(ligne, ":")(0)) = "mail" Then
             ligne = Replace(ligne, "'", "")
             objContact.Email1Address = Trim(Split(ligne, ":")(1))
            End If
            If Trim(Split(ligne, ":")(0)) = "telephoneNumber" Then
             objContact.OtherTelephoneNumber = Trim(Split(ligne, ":")(1))
            End If
            If Trim(Split(ligne, ":")(0)) = "postalCode" Then
             objContact.OtherAddressPostalCode = Trim(Split(ligne, ":")(1))
            End If
            If Trim(Split(ligne, ":")(0)) = "street" Then
             objContact.OtherAddressStreet = Trim(Split(ligne, ":")(1))
            End If
            If Trim(Split(ligne, ":")(0)) = "cn" Then
             objContact.FullName = Trim(Split(ligne, ":")(1))
             Dim prenom
             Dim nom
             nom = objContact.FirstName
             prenom = objContact.LastName
             objContact.FirstName = prenom
             objContact.LastName = nom
            End If
        End If
     Next
    objContact.Save
    
End Sub

Function traitementligne(s As String)
traitementligne = s

Dim aDecoder As String
  If InStr(s, "::") Then
    aDecoder = Trim(Split(s, "::")(1))
    aDecoder = Base64Decode(aDecoder)
    traitementligne = Trim(Split(s, "::")(0)) & ":" & aDecoder
  End If
End Function

' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

Function findcont(ct As String)
    ct = Replace(ct, "'", "")
    Dim olApp As Outlook.Application
    Dim dossierContacts As Outlook.MAPIFolder
    Dim Contact As Outlook.ContactItem
 
    Set olApp = New Outlook.Application
    Set dossierContacts = olApp.GetNamespace("MAPI"). _
        GetDefaultFolder(olFolderContacts)
 
    'Recherche le contact dont le nom est saisi dans la cellule A1
    Set Contact = dossierContacts.Items.Find _
        ("[Email1Address] = '" & ct & "'")
    If Not Contact Is Nothing Then
        findcont = True
    Else
        findcont = False
    End If
End Function




'   Char. number range  |        UTF-8 octet sequence
'      (hexadecimal)    |              (binary)
'   --------------------+---------------------------------------------
'   0000 0000-0000 007F | 0xxxxxxx
'   0000 0080-0000 07FF | 110xxxxx 10xxxxxx
'   0000 0800-0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx
'   0001 0000-0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
Public Function Encode_UTF8(astr)
    Dim c
    Dim n
    Dim utftext
     
    utftext = ""
    n = 1
    Do While n <= Len(astr)
        c = AscW(Mid(astr, n, 1))
        If c < 128 Then
            utftext = utftext + Chr(c)
        ElseIf ((c >= 128) And (c < 2048)) Then
            utftext = utftext + Chr(((c \ 64) Or 192))
            utftext = utftext + Chr(((c And 63) Or 128))
        ElseIf ((c >= 2048) And (c < 65536)) Then
            utftext = utftext + Chr(((c \ 4096) Or 224))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        Else ' c >= 65536
            utftext = utftext + Chr(((c \ 262144) Or 240))
            utftext = utftext + Chr(((((c \ 4096) And 63)) Or 128))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        End If
        n = n + 1
    Loop
    Encode_UTF8 = utftext
End Function
 
'   Char. number range  |        UTF-8 octet sequence
'      (hexadecimal)    |              (binary)
'   --------------------+---------------------------------------------
'   0000 0000-0000 007F | 0xxxxxxx
'   0000 0080-0000 07FF | 110xxxxx 10xxxxxx
'   0000 0800-0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx
'   0001 0000-0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
Public Function Decode_UTF8(astr)
    Dim c0, c1, c2, c3
    Dim n
    Dim unitext
     
    If isUTF8(astr) = False Then
        Decode_UTF8 = astr
        Exit Function
    End If
     
    unitext = ""
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If
         
        If (c0 And 240) = 240 And (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 240) * 65536 + (c1 - 128) * 4096) + (c2 - 128) * 64 + (c3 - 128)
            n = n + 4
        ElseIf (c0 And 224) = 224 And (c1 And 128) = 128 And (c2 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 224) * 4096 + (c1 - 128) * 64 + (c2 - 128))
            n = n + 3
        ElseIf (c0 And 192) = 192 And (c1 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 192) * 64 + (c1 - 128))
            n = n + 2
        ElseIf (c0 And 128) = 128 Then
            unitext = unitext + ChrW(c0 And 127)
            n = n + 1
        Else ' c0 < 128
            unitext = unitext + ChrW(c0)
            n = n + 1
        End If
    Loop
 
    Decode_UTF8 = unitext
End Function
 
'   Char. number range  |        UTF-8 octet sequence
'      (hexadecimal)    |              (binary)
'   --------------------+---------------------------------------------
'   0000 0000-0000 007F | 0xxxxxxx
'   0000 0080-0000 07FF | 110xxxxx 10xxxxxx
'   0000 0800-0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx
'   0001 0000-0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
Public Function isUTF8(astr)
    Dim c0, c1, c2, c3
    Dim n
     
    isUTF8 = True
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If
         
        If (c0 And 240) = 240 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
                n = n + 4
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 224) = 224 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 Then
                n = n + 3
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 192) = 192 Then
            If (c1 And 128) = 128 Then
                n = n + 2
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 128) = 0 Then
            n = n + 1
        Else
            isUTF8 = False
            Exit Function
        End If
    Loop
End Function



