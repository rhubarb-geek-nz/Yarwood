REM Copyright (c) 2024 Roger Brown.
REM Licensed under the MIT License.

Imports System.Reflection
Imports RhubarbGeekNzYarwood

Module Program
    Sub Main(args As String())
        Dim helloWorld = CreateObject("RhubarbGeekNz.AtYourService")
        Dim clientSecurity As DIClientSecurity = CreateObject("RhubarbGeekNz.ClientSecurity")
        clientSecurity.SetBlanket(helloWorld, UInt32.MaxValue, 0, Nothing, 4, 3, Nothing, 0)
        For hint As Int32 = 1 to 5
            Console.WriteLine(helloWorld.GetType().InvokeMember("GetMessage", BindingFlags.Public + BindingFlags.Instance + BindingFlags.InvokeMethod, Nothing, helloWorld, New Object() {hint}))
        Next
    End Sub
End Module
