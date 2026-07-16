Imports System
Imports System.Data
Imports MySql.Data.MySqlClient

Module Program
    Sub Main()
        Try
            Dim conStr As String = "Server=localhost;Database=hondadb;Uid=root;Pwd=1234;"
            Using con As New MySqlConnection(conStr)
                con.Open()
                
                Console.WriteLine("--- DB Snapshot ---")
                Dim queries = New String() {"departmentmaster", "department_list", "interchangeability", "type_details"}
                For Each tbl In queries
                    Try
                        Dim c As New MySqlCommand($"SELECT COUNT(*) FROM {tbl}", con)
                        Console.WriteLine($"{tbl} count: " & c.ExecuteScalar())
                    Catch ex As Exception
                        Console.WriteLine($"{tbl} error: " & ex.Message)
                    End Try
                Next
                
                Console.WriteLine("--- Department Master Data ---")
                Try
                    Dim c As New MySqlCommand($"SELECT DepartmentName FROM departmentmaster", con)
                    Using reader = c.ExecuteReader()
                        While reader.Read()
                            Console.WriteLine("'" & reader.GetString(0) & "'")
                        End While
                    End Using
                Catch ex As Exception
                End Try

            End Using
        Catch ex As Exception
            Console.WriteLine("Global Error: " & ex.Message)
        End Try
    End Sub
End Module
