﻿Imports System.Windows.Forms
Public Interface ICndaOutlookEvents
    Event XmlFileChangeEvent(ByVal xmlFilename As String,
                                   ByRef objList As CheckedListBox.ObjectCollection)
    Event SendEmailsEvent(ByRef objList As CheckedListBox.CheckedItemCollection,
                          ByRef count As Integer)
    Event PptFileChangeEvent(ByVal pptFilename As String)
End Interface
