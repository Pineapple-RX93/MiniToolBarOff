'Imports Microsoft.Win32
Imports System.Drawing
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Inventor

Namespace MiniToolBarOff
  <GuidAttribute("96951a56-37aa-436b-907e-21f6f046469f")> _
  Public Class StandardAddInServer
    Implements Inventor.ApplicationAddInServer
    ' Public Sub New()
    ' End Sub

#Region "Data"
    Private WithEvents oUserInputEvents As UserInputEvents
#End Region

#Region "ApplicationAddInServer Members"

    Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean
                        ) Implements Inventor.ApplicationAddInServer.Activate

      ' This method is called by Inventor when it loads the AddIn.
      ' The AddInSiteObject provides access to the Inventor Application object.
      ' The FirstTime flag indicates if the AddIn is loaded for the first time.
      ' Initialize AddIn members.
      AddinGlobal.IVApp = addInSiteObject.Application()

      Try
        AddinGlobal.GetAddinClassId(Me.GetType())
        oUserInputEvents = AddinGlobal.IVApp.CommandManager.UserInputEvents
      Catch ex As Exception
        System.Windows.Forms.MessageBox.Show(ex.ToString())
      End Try

    End Sub

    Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

      ' This method is called by Inventor when the AddIn is unloaded.
      ' The AddIn will be unloaded either manually by the user or
      ' when the Inventor session is terminated.

      ' TODO:  Add ApplicationAddInServer.Deactivate implementation
      If oUserInputEvents IsNot Nothing Then
        Marshal.FinalReleaseComObject(oUserInputEvents)
      End If
      If AddinGlobal.IVApp IsNot Nothing Then
        Marshal.FinalReleaseComObject(AddinGlobal.IVApp)
      End If

    End Sub

    Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

      ' This property is provided to allow the AddIn to expose an API 
      ' of its own to other programs. Typically, this  would be done by
      ' implementing the AddIn's API interface in a class and returning 
      ' that class object through this property.

      Get
        Return Nothing
      End Get

    End Property

    Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
      ' Note:this method is now obsolete, you should use the 
      ' ControlDefinition functionality for implementing commands.

    End Sub

    Private Sub oUserInputEvents_OnContextualMiniToolbar(ByVal SelectedEntities As ObjectsEnumerator, ByVal DisplayedCommands As NameValueMap,
                                                         ByVal AdditionalInfo As NameValueMap) Handles oUserInputEvents.OnContextualMiniToolbar
      DisplayedCommands.Clear()

    End Sub

#End Region

  End Class
End Namespace