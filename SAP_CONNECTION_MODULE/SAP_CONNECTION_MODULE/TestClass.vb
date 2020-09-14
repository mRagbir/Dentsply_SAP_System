Imports System.ValueTuple

Public Class TestClass



    Public Shared Function TryTupleAsArg(ByVal args As (String, String, String, Integer))

        Dim tup = (TestStuff1:=args.Item1, TestStuff2:=args.Item2, TestStuff3:=args.Item3, TestStuff4:=args.Item4)
        Return tup


    End Function


End Class
