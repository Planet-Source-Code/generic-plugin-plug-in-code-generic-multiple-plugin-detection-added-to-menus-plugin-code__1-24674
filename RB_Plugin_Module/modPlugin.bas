Attribute VB_Name = "modPlugin"
Option Explicit



Public Function AddPlugins(FormX As Object, FileList As FileListBox)

' This generic function will look for all plugins in a spesified directory.
' It will then query the plugin for identification and add the plugin
' to the main form.


Dim iIndex As Integer
Dim objTemp As Object
Dim sTemp As String
Dim sPlugin As String

'Fist, we need to get a list of all the exe (Plugin prefix) files in the directory
FileList.Pattern = "Plugin*.exe"
FileList.Refresh


'Now, we loop through all the plugin files and add them to the menus.
' In addition to this, we call a common function on the plugins that
' Identifies the plugins for us.
Dim ii As Integer
For ii = 0 To FileList.ListCount - 1

  sPlugin = Mid(FileList.List(ii), 1, Len(FileList.List(ii)) - 4) & ".clsPluginInterface"
  Set objTemp = CreateObject(sPlugin)
  sTemp = objTemp.Identify ' Run the function on the plugin to get the identification
  'add the plugin to the form's menus.
  iIndex = AddMenu(FormX, sTemp, sPlugin)
  Set objTemp = Nothing

Next ii

End Function
Public Sub RunPlugin(sPlugin As String, FormX As Form)

On Error GoTo Error_H

    'Declare a clean object to use
    Dim objPlugIn As Object
    Dim strResponse As String
    
    ' Run the Plugin
    'Set objPlugIn = CreateObject(Combo1.Text)
    Set objPlugIn = CreateObject(sPlugin)
    strResponse = objPlugIn.Run(FormX)
    
    'if the plug-in returns an error, let us know
    If strResponse <> vbNullString Then
        MsgBox strResponse
    End If
    
Exit Sub

Error_H:

MsgBox sPlugin & " - Error executing the plugin" & vbCrLf & Err.Description

End Sub


Public Function AddMenu(FormX As Object, sCaption As String, sTag As String) As Integer

Dim iIndex As Integer

iIndex = FormX.mnuPlugin.Count ' Get the position (Index) of where the plugin must go.

With FormX
  Load .mnuPlugin(iIndex)
  .mnuPlugin(iIndex).Caption = sCaption ' sCaption we got from the "Identify" function on the plugin
  .mnuPlugin(iIndex).Visible = True
  .mnuPlugin(iIndex).Enabled = True
  .mnuPlugin(iIndex).Tag = sTag ' We store the interface to the plugin in here, to later use it on the event of a menu click
End With

End Function

