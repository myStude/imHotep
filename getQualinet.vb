Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
Imports MySql.Data.MySqlClient
Imports System.Threading

''' <summary>
''' busca dados do qualinet
''' </summary>
''' <remarks></remarks>
Public Class getQualinet
    Inherits objConnection

    Sub writaGrid(T As String)

        Using qConn As New MySqlConnection(qualinetConnection("paq_last"))
            qConn.Open()
            Dim qComm As New MySqlCommand(qQueryStr, qConn)
            Using ReadRows As MySqlDataReader = qComm.ExecuteReader()

            End Using
        End Using


    End Sub

    Function qQueryStr() As String



        Return Nothing
    End Function

End Class
