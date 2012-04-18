imports Extensibility
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Visio

#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the $SAFEOBJNAME$Setup project, 
' right click the project in the Solution Explorer, then choose install.
#End Region

<GuidAttribute("4ABD85A3-13A2-43A1-A6A9-BC80E57C9FAC"), ProgIdAttribute("ExportPDF.Connect")> _
Public Class Connect
	
	Implements Extensibility.IDTExtensibility2

	Dim applicationObject as Object
    Dim addInInstance As Object
    Dim visio As Microsoft.Office.Interop.Visio.Application

	Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
    End Sub
	
	Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
	End Sub
	
    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete
        Try
            AddHandler visio.BeforeDocumentSave, AddressOf Create_PDF
            'Dim fileBar As Microsoft.Office.Core.CommandBar
            'fileBar = visio.CommandBars("File")
            'With fileBar.Controls.Add(Type:=Microsoft.Office.Core.MsoControlType.msoControlButton)
            '  .Tag = "File.ConvertToPdfandSave"
            '   .Caption = "To PDF and Save"
            'End With
        Catch ex As Exception
            MsgBox("Error registering handler in ExportPDF")
            ' MsgBox("Error registering CommandButton in ExportPDF")
        End Try
    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection
    End Sub

    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As Extensibility.ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection
        applicationObject = application
        addInInstance = addInInst
        Try
            ' Do an explicit cast to the Visio Application object so it is
            ' clear that there is a type change in this statement.      
            visio = CType(application, _
                 Microsoft.Office.Interop.Visio.Application)
            ' Show a message box with information about Visio using
            ' properties of the  Visio Application object.


        Catch err As COMException
            MsgBox("Exception in OnConnection: " & _
                err.Message, , "titulo")
        Catch err As InvalidCastException
            MsgBox("Exception in OnConnection: " & _
                err.Message, , "titulo")
        End Try

    End Sub

    Public Sub Create_PDF(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Try
            Dim filename As String

            filename = doc.FullName

            Dim i, pos, l As Integer
            pos = 1
            l = Len(filename)

            i = InStr(1, filename, ".")
            While i > 0
                pos = i
                i = InStr(pos + 1, filename, ".")
            End While

            filename = Mid(filename, 1, pos) & "pdf"

            doc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, filename, VisDocExIntent.visDocExIntentPrint, VisPrintOutRange.visPrintAll)
        Catch
            MsgBox("Error converting document")
        End Try
    End Sub

End Class
