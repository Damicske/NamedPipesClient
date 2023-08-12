Imports System.ComponentModel
Imports System.IO
Imports System.IO.Pipes
Imports System.Security.Principal
Imports System.Text
Imports System.Threading

Public Class Form1
    Private pipeClient As NamedPipeClientStream
    Private ClosingForm As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Console.WriteLine("*** Named pipe client stream with impersonation example ***")
        Dim client As New Thread(AddressOf ClientThread)
        client.Start()
    End Sub

    Private Sub ClientThread()
        While Not ClosingForm
            If pipeClient Is Nothing OrElse Not pipeClient.IsConnected Then
                pipeClient = New NamedPipeClientStream(".", "testpipe", PipeDirection.InOut, PipeOptions.None, TokenImpersonationLevel.Impersonation)

                Console.WriteLine("Connecting to server...")
                pipeClient.Connect()

                Dim ss As New StreamString(pipeClient)
                ' Validate the server's signature string.
                If ss.ReadString() = "I am the one true server!" Then
                    ' The client security token is sent with the first write.
                    ' Send the name of the file whose contents are returned
                    ' by the server.
                    ss.WriteString("getdata")

                    ' Print the file to the screen.
                    Debug.WriteLine(ss.ReadString())
                Else
                    Console.WriteLine("Server could not be verified.")
                End If
                'If pipeClient.
            End If
            Thread.Sleep(500)
        End While
        pipeClient.Close()
    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ClosingForm = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not pipeClient.IsConnected Then Exit Sub

        Try
            Dim ss As New StreamString(pipeClient)
            ' Validate the server's signature string.
            If ss.ReadString() = "I am the one true server!" Then
                ' The client security token is sent with the first write.
                ' Send the name of the file whose contents are returned
                ' by the server.
                ss.WriteString("getdata")

                ' Print the file to the screen.
                ListBox1.Items.Add(ss.ReadString())
            Else
                Console.WriteLine("Server could not be verified.")
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Class

''' <summary>
''' Defines the data protocol for reading and writing strings on our stream
''' </summary>
Public Class StreamString
    Private ReadOnly IoStream As Stream
    Private ReadOnly StreamEncoding As UnicodeEncoding

    Public Sub New(ioStream As Stream)
        Me.ioStream = ioStream
        streamEncoding = New UnicodeEncoding(False, False)
    End Sub

    ''' <summary>
    ''' Reads the stream for a string
    ''' </summary>
    Public Function ReadString() As String
        Try
            Dim Len As Integer = CType(IoStream.ReadByte(), Integer) * 256
            Len += CType(IoStream.ReadByte(), Integer)
            Dim inBuffer As Array = Array.CreateInstance(GetType(Byte), Len)
            IoStream.Read(inBuffer, 0, Len)
            Return StreamEncoding.GetString(CType(inBuffer, Byte()))
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Writes the outString to the given stream
    ''' </summary>
    ''' <param name="outString">String to write to stream</param>
    ''' <returns>Buffer length</returns>
    Public Function WriteString(outString As String) As Integer
        Dim outBuffer As Byte() = streamEncoding.GetBytes(outString)
        Dim len As Integer = outBuffer.Length
        If len > UShort.MaxValue Then len = CType(UShort.MaxValue, Integer)

        IoStream.WriteByte(CType(len \ 256, Byte))
        ioStream.WriteByte(CType(len And 255, Byte))
        ioStream.Write(outBuffer, 0, outBuffer.Length)
        ioStream.Flush()

        Return outBuffer.Length + 2
    End Function
End Class
