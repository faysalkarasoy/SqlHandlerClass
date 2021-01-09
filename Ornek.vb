Public Class OrnekSinif
    
    Sub Ornekler()
        Dim cs As String = "Data Source=.\MSSQLSERVER2017;Initial Catalog=ExampleDatabase;Persist Security Info=True;User ID=SqlUserName;Password=123456789;MultipleActiveResultSets=True"

        Dim sh As New SqlHandlerClass With {
            .ConnectionString = cs
        }
        '//Or
        Dim sh As New SqlHandlerClass(cs)
        '//Or
        Dim sh As New SqlHandlerClass("Data Source=.\MSSQLSERVER2017;Initial Catalog=ExampleDatabase;Persist Security Info=True;User ID=SqlUserName;Password=123456789;MultipleActiveResultSets=True")


        '// Examples
        '///////////////////////////////////////////////////////////////////////////
        '--------------------------------------------------------------------------


        'Execute
        '//////////////////////////////////////////////
        sh.ExeCuteNonQuery("UPDATE TableName SET [UserName]='test' WHERE Id=1")
        '//Or
        sh.ExeCuteNonQuery("UPDATE TableName SET [UserName]=@u WHERE Id=1", "@u", "test")
        '//Or

        Dim pList As New List(Of ISqlParameters)
        pList.Add(New ISqlParameters With {.Name = "@u", .Value = "test"})
        pList.Add(New ISqlParameters With {.Name = "@i", .Value = 1})
        sh.ExeCuteNonQuery("UPDATE TableName SET [UserName]=@u WHERE Id=i", pList)
        'Execute
        '//////////////////////////////////////////////


    End Sub
End Class
