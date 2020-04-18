Imports System.Windows.Forms
Public Interface ICndaOutlookEvents
    Event XmlFileChangeEvent(ByVal xmlFilename As String,
                                   ByRef objList As CheckedListBox.ObjectCollection)
    Event EmailFolderChangeEvent(ByRef emailFolder As Outlook.Folder)
    Event SendEmailsEvent(ByRef objList As CheckedListBox.CheckedItemCollection,
                                 ByRef count As Integer, ByVal pptFilename As String, ByVal genPdf As Boolean)
End Interface
