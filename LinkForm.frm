' De versie van VBA
VERSION 5.00

' Het begin van het formulier
' Hier staan alle eigenschappen van het formulier
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinkForm
    Caption         =   "Bulk Hyperlinker"
    ClientHeight    =   2050
    ClientLeft      =   110
    ClientTop       =   450
    ClientWidth     =   4580
    OleObjectBlob   =   "LinkForm.frx":0000
    StartUpPosition =   1  'CenterOwner
End

' Extra informatie over het formulier
Attribute VB_Name = "LinkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Deze functie wordt voor ons gegenereerd.
' De functie wordt aangeroepen als er op de annuleren knop wordt gedrukt.
Private Sub CancelButton_Click()
    ' Sluit het formulier af.
    Unload Me
' Het einde van de functie
End Sub

' Deze functie wordt aangeroepen als er op de ok knop wordt gedrukt.
Private Sub ConfirmButton_Click()
   ' Selecteer alle cellen
    SelectedRange = Selection.Address

    ' Ga langs alle cellen in de selectie
    For Each Cell In Range(SelectedRange)
        ' Haal de hyperlink op uit het tekstvak
        Hyperlink = UrlField.Value
                
        ' Controleer of de inhoud van de cell moet worden toegevoegd
        If AppendBox.Value = True Then
            ' Ja, voeg de inhoud toe
            Hyperlink = Hyperlink + Cell.Value
        ' Einde van de controle
        End If
        
        ' Stel de hyperlink in
        ActiveSheet.Hyperlinks.Add Range(Cell.Address), Address:=Hyperlink
    ' Ga door met de volgende cell
    Next Cell
End Sub
