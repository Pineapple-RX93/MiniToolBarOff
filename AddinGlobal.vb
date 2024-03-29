﻿Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Inventor

'Namespace InventorNetAddinVB1
Public Class AddinGlobal
  Public Shared IVApp As Inventor.Application
  Public Shared RibbonPanelId As String
  Public Shared RibbonPanel As RibbonPanel
  Public Shared ButtonList As New List(Of InventorButton)()
  Private Shared mClassId As String

  Public Shared Property ClassId() As String
    Get
      If Not String.IsNullOrEmpty(mClassId) Then
        Return AddinGlobal.mClassId
      Else
        Throw New System.Exception("The addin server class id hasn't been gotten yet!")
      End If
    End Get
    Set(value As String)
      AddinGlobal.mClassId = value
    End Set
  End Property

  Public Shared Sub GetAddinClassId(t As Type)
    Dim guidAtt As GuidAttribute = DirectCast(GuidAttribute.GetCustomAttribute(t, GetType(GuidAttribute)), GuidAttribute)
    mClassId = "{" & guidAtt.Value & "}"
  End Sub
End Class
'End Namespace


