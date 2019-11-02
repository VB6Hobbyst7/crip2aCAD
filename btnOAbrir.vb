#Region "Imported Namespaces"
Imports System
Imports System.Collections.Generic
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
'Imports UR2ACAD2016.btnWeb2aCAD
Imports crip2aCAD
#End Region

'Namespace UCRevit2018
<Transaction(TransactionMode.Manual)>
Public Class btnOAbrir
    Implements Autodesk.Revit.UI.IExternalCommand
    ''
    ''' <summary>
    ''' The one and only method required by the IExternalCommand interface, the main entry point for every external command.
    ''' </summary>
    ''' <param name="commandData">Input argument providing access to the Revit application, its documents and their properties.</param>
    ''' <param name="message">Return argument to display a message to the user in case of error if Result is not Succeeded.</param>
    ''' <param name="elements">Return argument to highlight elements on the graphics screen if Result is not Succeeded.</param>
    ''' <returns>Cancelled, Failed or Succeeded Result code.</returns>
    Public Function Execute(
      ByVal commandData As ExternalCommandData,
      ByRef message As String,
      ByVal elements As ElementSet) _
    As Result Implements IExternalCommand.Execute
        ' Llenamos las variables que necesitemos.
        clsCR.crAppUI = commandData.Application
        ''
        'TODO: Add your code here
        Dim resultado As Result = Result.Succeeded
        '





        '
Final:
        ' Vaciar las variables innecesarias.
        '
        Return resultado
    End Function
End Class
'End Namespace
