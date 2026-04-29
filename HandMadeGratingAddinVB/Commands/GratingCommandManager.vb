'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' GratingCommandManager: Registers and manages all add-in commands.
' Keeps command setup separate from the add-in server lifecycle.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports Inventor

''' <summary>
''' Central manager for all add-in command definitions and ribbon placement.
''' </summary>
Public Class GratingCommandManager

    Private _gratingCommand As GratingCommand
    Private _dockPanel As DockableWindowManager

    ''' <summary>
    ''' Creates all command definitions. Call every time the add-in activates.
    ''' </summary>
    Public Sub Initialize(app As Application, clientId As String)
        Trace.TraceInformation(": HandMadeGratingAddinVB: Registering commands...")

        ' Initialize the dockable panel (non-fatal if it fails)
        _dockPanel = New DockableWindowManager()
        _dockPanel.Initialize(app, clientId)

        _gratingCommand = New GratingCommand()
        _gratingCommand.DockPanel = _dockPanel
        _gratingCommand.CreateDefinition(app, clientId)
    End Sub

    ''' <summary>
    ''' Adds commands to the Inventor ribbon. Only call when firstTime is True.
    ''' </summary>
    Public Sub CreateUserInterface(app As Application, clientId As String)
        Trace.TraceInformation(": HandMadeGratingAddinVB: Creating user interface...")
        _gratingCommand.AddToRibbon(app, clientId)
    End Sub

    ''' <summary>
    ''' Cleans up event handlers. Call during Deactivate.
    ''' </summary>
    Public Sub Cleanup()
        If _gratingCommand IsNot Nothing Then
            _gratingCommand.Cleanup()
            _gratingCommand = Nothing
        End If
        If _dockPanel IsNot Nothing Then
            _dockPanel.Cleanup()
            _dockPanel = Nothing
        End If
    End Sub

End Class
