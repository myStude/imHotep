Imports System.IO

Public Class writerController

    Public Sub write(ByVal dtHora As String, ByVal strTexto As String)

        'Dim NOMELOG As String = "_" & Environment.MachineName & "_" & Format(Now, "yyyyMMdd")
        'Dim fileFolder As String = "\\npoasv0032\PrivateShare\LogGerenciadores\IMhotep" & NOMELOG & ".txt"

        'Dim log As StreamWriter
        'Dim texto As String = ""
        'Dim fluxoTexto As IO.StreamReader
        'Dim linhaTexto As String
        'If IO.File.Exists(fileFolder) Then
        '    fluxoTexto = New IO.StreamReader(fileFolder)
        '    linhaTexto = fluxoTexto.ReadLine

        '    While linhaTexto <> Nothing
        '        texto = "# " & texto & linhaTexto & vbNewLine
        '        linhaTexto = fluxoTexto.ReadLine
        '    End While
        '    fluxoTexto.Close()
        'End If
        'texto = Environment.MachineName & " - " & dtHora & " - " & strTexto & vbNewLine
        'log = New StreamWriter(fileFolder)
        'log.Write(texto)
        'log.Close()

    End Sub
End Class
