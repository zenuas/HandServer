Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Threading.Tasks


Public Class Main

    Public Shared Sub Main(args() As String)

        Dim self As New Main
        self.Run(CommandLineParser.Parse(self, args))
    End Sub

    Public Overridable Sub Run(args() As String)

        Dim listener As New TcpListener(IPAddress.Any, Me.LocalPort)
        listener.Start()

        Dim do_accept = Task.Run(
          Async Function()
              Do While True

                  Dim socket = Await listener.AcceptSocketAsync
                  Dim stdout = Console.OpenStandardOutput
                  Dim stdin = Console.OpenStandardInput

                  Dim done_receive = False
                  Dim do_receive = Task.Run(
                      Sub()

                          Dim buffer(1024) As Byte
                          Do While True

                              Dim count = socket.Receive(buffer)
                              If count <= 0 Then

                                  done_receive = True
                                  Return
                              End If
                              stdout.Write(buffer, 0, count)
                              stdout.Flush()
                          Loop
                      End Sub)

                  Dim done_send = False
                  Dim do_send = Task.Run(
                      Sub()

                          Dim buffer(1024) As Byte
                          Do While True

                              Dim count = stdin.Read(buffer, 0, buffer.Length)
                              If count <= 0 Then

                                  done_send = True
                                  Return
                              End If
                              socket.Send(buffer, count, SocketFlags.None)
                          Loop
                      End Sub)

                  Do While Not done_receive OrElse Not done_send

                      Thread.Sleep(100)
                  Loop
              Loop
          End Function)

        Do While True

            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite)
        Loop

    End Sub

    <Argument("P"c, , "local-port")>
    Public Overridable Property LocalPort As Integer = 0

End Class
