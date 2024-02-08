
Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor

'Namespace InventorNetAddinVB1
''' <summary>
''' The class wrapps up Inventor Button creation stuffs and is easy to use.
''' No need to derive. Create an instance using either constructor and assign the Action.
''' </summary>
Public Class InventorButton
#Region "Fields & Properties"

  Private mButtonDef As ButtonDefinition
  Public Property ButtonDef() As ButtonDefinition
    Get
      Return mButtonDef
    End Get
    Set(ByVal value As ButtonDefinition)
      mButtonDef = value
    End Set
  End Property

#End Region

#Region "Constructors"

  ''' <summary>
  ''' The most comprehensive signature.
  ''' </summary>
  Public Sub New(ByVal displayName As String, ByVal internalName As String, ByVal description As String, ByVal tooltip As String, ByVal standardIcon As Icon, ByVal largeIcon As Icon, _
   ByVal commandType As CommandTypesEnum, ByVal buttonDisplayType As ButtonDisplayEnum)
    Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, standardIcon, _
     largeIcon, commandType, buttonDisplayType)
  End Sub

  ''' <summary>
  ''' The signature does not care about Command Type (always editing) and Button Display (always with text).
  ''' </summary>
  Public Sub New(ByVal displayName As String, ByVal internalName As String, ByVal description As String, ByVal tooltip As String, ByVal standardIcon As Icon, ByVal largeIcon As Icon)
    Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, Nothing, _
     Nothing, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
  End Sub

  ''' <summary>
  ''' The signature does not care about icons.
  ''' </summary>
  Public Sub New(ByVal displayName As String, ByVal internalName As String, ByVal description As String, ByVal tooltip As String, ByVal commandType As CommandTypesEnum, ByVal buttonDisplayType As ButtonDisplayEnum)
    Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, Nothing, _
     Nothing, commandType, buttonDisplayType)
  End Sub

  ''' <summary>
  ''' This signature only cares about display name and icons.
  ''' </summary>
  ''' <param name="displayName"></param>
  ''' <param name="standardIcon"></param>
  ''' <param name="largeIcon"></param>
  Public Sub New(ByVal displayName As String, ByVal standardIcon As Icon, ByVal largeIcon As Icon)
    Create(displayName, displayName, displayName, displayName, AddinGlobal.ClassId, standardIcon, _
     largeIcon, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
  End Sub

  ''' <summary>
  ''' The simplest signature, which can be good for prototyping.
  ''' </summary>
  Public Sub New(ByVal displayName As String)
    Create(displayName, displayName, displayName, displayName, AddinGlobal.ClassId, Nothing, _
     Nothing, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
  End Sub

  ''' <summary>
  ''' The helper method for constructors to call to avoid duplicate code.
  ''' </summary>
  Public Sub Create(ByVal displayName As String, ByVal internalName As String, ByVal description As String, ByVal tooltip As String, ByVal clientId As String, ByVal standardIcon As Icon, _
   ByVal largeIcon As Icon, ByVal commandType As CommandTypesEnum, ByVal buttonDisplayType As ButtonDisplayEnum)
    If String.IsNullOrEmpty(clientId) Then
      clientId = AddinGlobal.ClassId
    End If

    Dim standardIconIPictureDisp As stdole.IPictureDisp = Nothing
    Dim largeIconIPictureDisp As stdole.IPictureDisp = Nothing
    If standardIcon IsNot Nothing Then
      standardIconIPictureDisp = IconToPicture(standardIcon)
      largeIconIPictureDisp = IconToPicture(largeIcon)
    End If

    mButtonDef = AddinGlobal.IVApp.CommandManager.ControlDefinitions.AddButtonDefinition(displayName, internalName, commandType, clientId, description, tooltip, _
     standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType)

    mButtonDef.Enabled = True
    AddHandler mButtonDef.OnExecute, AddressOf ButtonDefinition_OnExecute

    DisplayText = True

    AddinGlobal.ButtonList.Add(Me)
  End Sub

#End Region

#Region "Behavior"

  Public Property DisplayBigIcon() As Boolean
    Get
      Return m_DisplayBigIcon
    End Get
    Set(ByVal value As Boolean)
      m_DisplayBigIcon = value
    End Set
  End Property
  Private m_DisplayBigIcon As Boolean

  Public Property DisplayText() As Boolean
    Get
      Return m_DisplayText
    End Get
    Set(ByVal value As Boolean)
      m_DisplayText = value
    End Set
  End Property
  Private m_DisplayText As Boolean

  Public Property InsertBeforeTarget() As Boolean
    Get
      Return m_InsertBeforeTarget
    End Get
    Set(ByVal value As Boolean)
      m_InsertBeforeTarget = value
    End Set
  End Property
  Private m_InsertBeforeTarget As Boolean

  Public Sub SetBehavior(ByVal displayBigIcon__1 As Boolean, ByVal displayText__2 As Boolean, ByVal insertBeforeTarget__3 As Boolean)
    DisplayBigIcon = displayBigIcon__1
    DisplayText = displayText__2
    InsertBeforeTarget = insertBeforeTarget__3
  End Sub

  Public Sub CopyBehaviorFrom(ByVal button As InventorButton)
    Me.DisplayBigIcon = button.DisplayBigIcon
    Me.DisplayText = button.DisplayText
    Me.InsertBeforeTarget = Me.InsertBeforeTarget
  End Sub

#End Region

#Region "Actions"

  ''' <summary>
  ''' The button callback method.
  ''' </summary>
  ''' <param name="context"></param>
  Private Sub ButtonDefinition_OnExecute(ByVal context As NameValueMap)
    If Execute IsNot Nothing Then
      Execute()
    Else
      MessageBox.Show("Nothing to execute.")
    End If
  End Sub

  ''' <summary>
  ''' The button action to be assigned from anywhere outside.
  ''' </summary>
  Public Execute As Action

#End Region

#Region "Image Converters"

  Public Shared Function ImageToPicture(ByVal image As Image) As stdole.IPictureDisp
    Return ImageConverter.ImageToPicture(image)
  End Function

  Public Shared Function IconToPicture(ByVal icon As Icon) As stdole.IPictureDisp
    Return ImageConverter.ImageToPicture(icon.ToBitmap())
  End Function

  Public Shared Function PictureToImage(ByVal picture As stdole.IPictureDisp) As Image
    Return ImageConverter.PictureToImage(picture)
  End Function

  Public Shared Function PictureToIcon(ByVal picture As stdole.IPictureDisp) As Icon
    Return ImageConverter.PictureToIcon(picture)
  End Function

  Private Class ImageConverter
    Inherits AxHost
    Public Sub New()
      MyBase.New(String.Empty)
    End Sub

    Public Shared Function ImageToPicture(ByVal image As Image) As stdole.IPictureDisp
      Return DirectCast(GetIPictureDispFromPicture(image), stdole.IPictureDisp)
    End Function

    Public Shared Function IconToPicture(ByVal icon As Icon) As stdole.IPictureDisp
      Return ImageToPicture(icon.ToBitmap())
    End Function

    Public Shared Function PictureToImage(ByVal picture As stdole.IPictureDisp) As Image
      Return GetPictureFromIPicture(picture)
    End Function

    Public Shared Function PictureToIcon(ByVal picture As stdole.IPictureDisp) As Icon
      Dim bitmap As New Bitmap(PictureToImage(picture))
      Return System.Drawing.Icon.FromHandle(bitmap.GetHicon())
    End Function
  End Class

#End Region

End Class
'End Namespace


