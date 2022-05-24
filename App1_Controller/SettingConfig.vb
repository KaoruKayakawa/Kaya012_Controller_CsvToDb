Imports System.Xml

Public Class SettingConfig
    Protected Shared _root As XmlNode

    Shared Sub New()
        Try
            Dim fn As String
            fn = System.Environment.CurrentDirectory + "\Setting.config"

            Dim doc As XmlDocument = New XmlDocument
            doc.Load(fn)

            _root = doc.SelectSingleNode("/configuration")
        Catch ex As Exception
            Throw New ApplicationException("Iregular Setting.config File.", ex)
        End Try
    End Sub

    Private Shared ReadOnly Property ExeApps As XmlNode
        Get
            Return _root.SelectSingleNode("./exe_apps")
        End Get
    End Property

    Public Shared ReadOnly Property ExeApps_MaxParallel As Integer
        Get
            Return Integer.Parse(ExeApps.Attributes("max_parallel").Value)
        End Get
    End Property

    Public Shared ReadOnly Property ExeApps_App As XmlNodeList
        Get
            Return ExeApps.SelectNodes("./app")
        End Get
    End Property

    Public Shared Function App_Name(app_node As XmlNode) As String
        Return app_node.Attributes("name").Value
    End Function

    Public Shared Function App_Path(app_node As XmlNode) As String
        Return app_node.Attributes("path").Value
    End Function

End Class
