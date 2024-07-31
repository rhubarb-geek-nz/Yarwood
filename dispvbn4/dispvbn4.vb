REM Copyright (c) 2024 Roger Brown.
REM Licensed under the MIT License.

Module Program
    Sub Main(args As String())
        Dim runningMan = GetObject(, "RhubarbGeekNz.RunningMan")
        Dim clientSecurity = CreateObject("RhubarbGeekNz.ClientSecurity")
        clientSecurity.SetBlanket(runningMan, -1, 0, nothing, 4, 3, nothing, 0)
        For hint = 1 to 5
            Console.WriteLine(runningMan.GetMessage(hint))
        Next
    End Sub
End Module
