'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' GratingAutomation: Automation entry point exposed to Inventor /
' Design Automation via the StandardAddInServer.Automation property.
'
' Phase 1: Startup verification only.
' Phase 2+: Parameter processing, geometry generation, etc.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Inventor

<ComVisible(True)>
Public Class GratingAutomation

    Private ReadOnly _inventorServer As InventorServer

    Public Sub New(inventorServer As InventorServer)
        _inventorServer = inventorServer
    End Sub

    ''' <summary>
    ''' Called by Design Automation with the active document.
    ''' </summary>
    Public Sub Run(doc As Document)
        LogTrace("Run called with " & doc.DisplayName)
        ' Phase 2+: implement grating parameter and geometry logic
    End Sub

    ''' <summary>
    ''' Called by Design Automation with additional arguments.
    ''' </summary>
    Public Sub RunWithArguments(doc As Document, map As NameValueMap)
        LogTrace("RunWithArguments called with " & doc.DisplayName)
        ' Phase 2+: implement parameter processing and model update logic
    End Sub

#Region "Logging Utilities"

    Private Shared Sub LogTrace(message As String)
        Trace.TraceInformation(message)
    End Sub

    Private Shared Sub LogTrace(format As String, ParamArray args() As Object)
        Trace.TraceInformation(format, args)
    End Sub

    Private Shared Sub LogError(message As String)
        Trace.TraceError(message)
    End Sub

    Private Shared Sub LogError(format As String, ParamArray args() As Object)
        Trace.TraceError(format, args)
    End Sub

#End Region

End Class
