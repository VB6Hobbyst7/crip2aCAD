Imports System
Imports System.Timers
Imports System.Text
Imports System.Drawing
Imports System.Security
Imports System.Security.Cryptography
Imports System.Xml
Imports System.Linq
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Media.Imaging
Imports System.IO
Partial Public Class clsCR
    ' Objectos para tareas
    Public Shared oTBorraTodo As Task = Nothing
    Public Shared oTRevit As Task = Nothing
    Public Shared oTXML As Task = Nothing
    Public Shared oTBAK As Task = Nothing
    Public Shared _registered As Boolean
    '
    'Public Shared WithEvents oTimer_Cierra As System.Timers.Timer   ' Timer para Cerrar documento actual    
    Public Shared WithEvents oTimer_Borra As System.Timers.Timer                  ' Timer para Borrar fichero de folderTemp   
    Public Shared WithEvents oTimer_Abre As System.Timers.Timer                  ' Timer para Borrar fichero de folderTemp
    '
    Private Shared Sub oTimer_Borra_Elapsed(sender As Object, e As ElapsedEventArgs) Handles oTimer_Borra.Elapsed
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        'Activador_Ficheros_TodosBorraDeTemp()
        If crPuedoborrar = False Then Exit Sub
        FicherosDesarrollo_BorraTempTASK()
    End Sub
    '
    '
    Public Shared Sub FicherosDesarrollo_BorraTempTASK()
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        ' Quitamos temporalmente el borrado de TEMP
        If crPuedoborrar = False Then Exit Sub
        Try
            'If (crip2aCAD.clsCR.oTBorraTodo Is Nothing) Or (crip2aCAD.clsCR.oTBorraTodo IsNot Nothing AndAlso crip2aCAD.clsCR.oTBorraTodo.Status <> System.Threading.Tasks.TaskStatus.Running) Then
            crip2aCADUC.clsCR.oTBorraTodo = System.Threading.Tasks.Task.Run(AddressOf crip2aCADUC.clsCR.FicherosDesarrollo_BorraTemp)
            'End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub
    ' Borrar los fichero Revit que usa el desarrollo {".rvt", ".rfa", ".rte", ".xml", ".xmlc", ".adsk", ".adsklib", ".crypto", ".bak"} de la carpeta temp.
    Public Shared Sub FicherosDesarrollo_BorraTemp()
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If crPuedoborrar = False Then Exit Sub
        'Dim folderTemp As String = IO.Path.GetTempPath
        Dim mascara As String = "*.*"
        Dim extensiones() As String = New String() {".rvt", ".rfa", ".rte", ".xml", ".xmlc", ".adsk", ".adsklib", ".tempcrypto", ".bak", ".jpg"}
        '
        For Each fiBorra As String In IO.Directory.GetFiles(crfolder_Temp, mascara, SearchOption.AllDirectories)
            ' Saltar los fichero que no sean Revit (proyecto o familia) y que no sean los que estamos usando.
            If extensiones.Contains(IO.Path.GetExtension(fiBorra)) = False Then Continue For
            If crFicherosEnUso.Contains(fiBorra) = True Then Continue For
            'If fiBorra = crfiAbrir Then Continue For
            'If crDicFicheros.ContainsKey(fiBorra) Then Continue For
            'If crDicFicheros.ContainsValue(fiBorra) Then Continue For
            '
            Try
                'If Fichero_EnUso(fiBorra) = False Then IO.File.Delete(fiBorra)
                IO.File.Delete(fiBorra)
                Fichero_Quita(fiBorra, borrar:=False)      ' También lo borra
            Catch ex As Exception
                Debug.Print(ex.ToString())
            End Try
        Next
        '
        ' Ahora los directorios vacíos.
        For Each dirBorra As String In IO.Directory.GetDirectories(crfolder_Temp)
            Try
                Dim arrFi As String() = IO.Directory.GetFiles(dirBorra, "*.*", SearchOption.TopDirectoryOnly)
                If arrFi.Count = 0 Then
                    Try
                        IO.Directory.Delete(dirBorra, True)
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                End If
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
        Next
        '
        ' Los del directorio Temp. Por si había alguno
        'For Each queExt As String In extensiones
        '    Dim queMasc As String = crfolder_TempRaiz + IO.Path.DirectorySeparatorChar + "*" + queExt
        '    Try
        '        Dim files As String() = IO.Directory.GetFiles(crfolder_TempRaiz, "*" + queExt, SearchOption.TopDirectoryOnly)
        '        If files IsNot Nothing And files.Count > 0 Then
        '            Kill(queMasc)
        '        End If
        '    Catch ex As Exception
        '        'Debug.Print(ex.ToString)
        '    End Try
        'Next
        crip2aCADUC.clsCR.oTBorraTodo = Nothing
    End Sub
    '
End Class
