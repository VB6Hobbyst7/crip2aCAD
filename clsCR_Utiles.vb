Imports System.Collections.Generic
Imports System.Diagnostics '– used for debug 
'Imports System.Windows.Forms
Imports System.IO '– used for reading folders 
Imports System.Windows.Media.Imaging '– used for bitmap images 
Imports adWin = Autodesk.Windows
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.DB

Partial Public Class clsCR
    Public Const idModify As String = "Modify"
    Public Const idInsert As String = "Insert"
    Public Shared Sub PonLog(ByVal quetexto As String, Optional ByVal borrar As Boolean = False)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        Dim queF As String = cr_ficherolog
        If borrar = True Then IO.File.Delete(queF)
        If quetexto.EndsWith(vbCrLf) = False Then quetexto &= vbCrLf
        IO.File.AppendAllText(queF, Date.Now.ToString & vbTab & quetexto)
    End Sub

    Public Shared Function Documento_EstaAbierto(oApp As Autodesk.Revit.ApplicationServices.Application, fulldoc As String, Optional cerrar As Boolean = False) As Autodesk.Revit.DB.Document
        ' Si no está activado, salir ****
        If activado = False Then Return Nothing : Exit Function
        ' ********************************
        Dim resultado As Autodesk.Revit.DB.Document = Nothing
        Dim oAppUI As New Autodesk.Revit.UI.UIApplication(oApp)
        '
        If oApp.Documents.Size = 0 Then
            Return resultado
            Exit Function
        End If

        For Each queDoc As Autodesk.Revit.DB.Document In oApp.Documents
            If queDoc.PathName = fulldoc Then
                If cerrar = False Then
                    resultado = queDoc
                Else
                    If oAppUI.ActiveUIDocument.Document.PathName <> queDoc.PathName Then
                        Try
                            queDoc.Close(True)
                            resultado = Nothing
                        Catch ex As Exception
                            resultado = Nothing
                        End Try
                    End If
                End If
                Exit For
            End If
        Next
        oAppUI = Nothing
        Return resultado
    End Function
    Public Shared Sub BorraTempAlSalir()
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        Try
            '/BorraTempRevit|G:\Temp\TEMPS
            Dim p As Process = New Process()
            'p.WaitForExit(1000 * 40)
            p.PriorityBoostEnabled = True
            Dim psi As ProcessStartInfo = New ProcessStartInfo(IO.Path.Combine(My.Application.Info.DirectoryPath, "task2aCAD.exe"))
            psi.UseShellExecute = True
            psi.Arguments = "/BorraTempRevit|G:\Temp\TEMPS"
            psi.WindowStyle = ProcessWindowStyle.Hidden
            'psi.CreateNoWindow = True
            p.StartInfo = psi
            p.Start()
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Function Fichero_EnUso(FileIn As String) As Boolean
        ' Si no está activado, salir ****
        If activado = False Then Return Nothing : Exit Function
        ' ********************************
        Dim resultado As Boolean = False
        Try
            Fichero_ReadOnly(FileIn, sololectura:=False)
            fsInput = New System.IO.FileStream(FileIn, FileMode.Open,
                                                      FileAccess.ReadWrite)
        Catch ex As Exception
            '' El fichero está en uso y no se puede escribir en él.
            resultado = True
        End Try
        '
        If fsInput IsNot Nothing Then fsInput.Close()
        fsInput = Nothing
        Return resultado
    End Function
    '
    Public Shared Sub Fichero_ReadOnly_Scripting(queFi As String, sololectura As Boolean)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If IO.File.Exists(queFi) = False Then Exit Sub
        '
        With CreateObject("scripting.filesystemobject")
            With .GetFile(queFi)
                If .Attributes And 1 And sololectura = True Then
                    .Attributes = .Attributes - 1
                ElseIf .Attributes And -1 And sololectura = False Then
                    .Attributes = .Atributes + 1
                End If
            End With
        End With
    End Sub
    '0: normal (no se establece ningun atributo)
    '1: solo lectura
    '2: oculto
    '4: sistema (lectura-escritura)
    '8: etiqueta de volumen de unidad (solo lectura)
    '16: carpeta o directorio (solo lectura)
    '32: archivo (lectura-escritura)
    '64: alias (atajo) (solo lectura)
    '128: (comprimido)

    'Public Shared Sub Fichero_QuitarSoloLecturaFileInfo(queFi As String)
    '    If IO.File.Exists(queFi) = False Then Exit Sub
    '    '
    '    Try
    '        Dim FileInfo As New FileInfo(queFi)
    '        FileInfo.Attributes = IO.FileAttributes.Normal
    '    Catch ex As Exception
    '        Debug.Print(ex.ToString)
    '    End Try
    'End Sub

    Public Shared Sub Directorio_Hidden(queDir As String, ocultar As Boolean)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If IO.Directory.Exists(queDir) = False Then Exit Sub
        Dim oDinf As IO.DirectoryInfo = New IO.DirectoryInfo(queDir)
        Dim attribute As System.IO.FileAttributes = Nothing
        If ocultar = True Then
            attribute = oDinf.Attributes Or FileAttributes.Hidden
        Else
            attribute = IO.FileAttributes.Normal
        End If
        oDinf.Attributes = attribute
        oDinf = Nothing
    End Sub
    Public Shared Sub Fichero_Hidden(queFic As String, ocultar As Boolean)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If IO.File.Exists(queFic) = False Then Exit Sub
        Dim oFinf As IO.FileInfo = New IO.FileInfo(queFic)
        Dim attribute As System.IO.FileAttributes = Nothing
        If ocultar = True Then
            attribute = oFinf.Attributes Or FileAttributes.Hidden
        Else
            attribute = IO.FileAttributes.Normal
        End If
        oFinf.Attributes = attribute
        oFinf = Nothing
    End Sub
    Public Shared Sub Fichero_ReadOnly(queFic As String, sololectura As Boolean)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If IO.File.Exists(queFic) = False Then Exit Sub
        Try
            Dim oFinf As IO.FileInfo = New IO.FileInfo(queFic)
            Dim attribute As System.IO.FileAttributes = Nothing
            If sololectura = True Then
                attribute = oFinf.Attributes Or FileAttributes.ReadOnly
            Else
                attribute = IO.FileAttributes.Normal
            End If
            oFinf.Attributes = attribute
            oFinf = Nothing
        Catch ex As Exception
            Debug.Print(ex.ToString)
            PonLog("ERROR " & ex.ToString)
        End Try
    End Sub
End Class