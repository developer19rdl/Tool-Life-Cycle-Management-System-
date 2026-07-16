Imports System.Windows
Imports System.Data

Public Class CycleManager

    Public Shared Function GetNaturalCycleNameForDate(dt As DateTime, mysql As MySQLClass) As String
        Try
            Dim gsVal = mysql.GetConfigValue("GreenCycleStart")
            Dim geVal = mysql.GetConfigValue("GreenCycleEnd")
            Dim ysVal = mysql.GetConfigValue("YellowCycleStart")
            Dim yeVal = mysql.GetConfigValue("YellowCycleEnd")

            If Not String.IsNullOrEmpty(gsVal) AndAlso Not String.IsNullOrEmpty(geVal) AndAlso
               Not String.IsNullOrEmpty(ysVal) AndAlso Not String.IsNullOrEmpty(yeVal) Then

                Dim greenStart As DateTime
                Dim greenEnd As DateTime
                Dim yellowStart As DateTime
                Dim yellowEnd As DateTime

                If DateTime.TryParse(gsVal, greenStart) AndAlso DateTime.TryParse(geVal, greenEnd) AndAlso
                   DateTime.TryParse(ysVal, yellowStart) AndAlso DateTime.TryParse(yeVal, yellowEnd) Then

                    If dt >= greenStart AndAlso dt <= greenEnd Then
                        Return $"{greenStart.ToString("MMM")}'{greenStart.ToString("yy")} Green Cycle"
                    End If
                    If dt >= yellowStart AndAlso dt <= yellowEnd Then
                        Return $"{yellowStart.ToString("MMM")}'{yellowStart.ToString("yy")} Yellow Cycle"
                    End If
                    
                    Dim startColor = mysql.GetConfigValue("CycleStartColor")
                    If startColor = "Yellow" Then
                        Return $"{yellowStart.ToString("MMM")}'{yellowStart.ToString("yy")} Yellow Cycle"
                    Else
                        Return $"{greenStart.ToString("MMM")}'{greenStart.ToString("yy")} Green Cycle"
                    End If
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("GetNaturalCycleNameForDate config read error: " & ex.Message)
        End Try

        If dt.Month >= 1 AndAlso dt.Month <= 6 Then
            Return $"Jan'{dt.ToString("yy")} Yellow Cycle"
        Else
            Return $"Jul'{dt.ToString("yy")} Green Cycle"
        End If
    End Function

    Public Shared Function PerformRollover(activeCycle As String, nextCycle As String, mysql As MySQLClass, currentUser As String) As Boolean
        Try
            mysql.LogCycleHistoryBulk(activeCycle)

            Dim dtTemp = mysql.ReadDatatable($"SELECT * FROM interchangeability WHERE CycleName = '{activeCycle.Replace("'", "''")}' AND Status = 'Temp issuance'")
            For Each row As DataRow In dtTemp.Rows
                Dim ctrlNo = row("ControlNo").ToString()
                Dim dept = row("Department").ToString()
                Dim instName = row("InstrumentName").ToString()
                Dim color = row("Color").ToString()
                Dim size = row("SizeandRange").ToString()

                Dim cat = mysql.GetItemCategory(ctrlNo)
                Dim remark = If(cat.Equals("Gauge", StringComparison.OrdinalIgnoreCase), "Gauge missing", "Instrument missing")

                mysql.InsertWOPRecord(ctrlNo, nextCycle, dept, instName, color, remark)
                mysql.InsertInterchangeRecord(nextCycle, ctrlNo, dept, instName, size, color, "WOP", remark)
            Next

            Dim dtPending = mysql.ReadDatatable($"SELECT * FROM interchangeability WHERE CycleName = '{activeCycle.Replace("'", "''")}' AND (Status = 'Pending' OR Status IS NULL OR Status = '')")
            For Each row As DataRow In dtPending.Rows
                Dim ctrlNo = row("ControlNo").ToString()
                Dim dept = row("Department").ToString()
                Dim instName = row("InstrumentName").ToString()
                Dim color = row("Color").ToString()

                Dim cycleIsYellow = activeCycle.Contains("Yellow")
                Dim cycleIsGreen = activeCycle.Contains("Green")
                Dim gaugeIsYellow = color.Equals("Yellow", StringComparison.OrdinalIgnoreCase)
                Dim gaugeIsGreen = color.Equals("Green", StringComparison.OrdinalIgnoreCase)
                Dim colorsMatch = (cycleIsYellow AndAlso gaugeIsYellow) OrElse (cycleIsGreen AndAlso gaugeIsGreen)

                Dim remark = If(colorsMatch, "Not Issued", "Not Recieved")
                mysql.InsertWOPRecord(ctrlNo, nextCycle, dept, instName, color, remark)
            Next

            Dim query = "INSERT INTO interchangeability (CycleName, ControlNo, Department, InstrumentName, SizeandRange, Color, Status, ActionDate, ActionTime, Remarks, ActionBy) " &
                        $"SELECT '{nextCycle.Replace("'", "''")}', ControlNo, Department, InstrumentName, SizeandRange, Color, " &
                        "CASE " &
                        "  WHEN Status = 'WOP' THEN 'WOP' " &
                        "  WHEN Status = 'Pending' OR Status IS NULL OR Status = '' THEN 'WOP' " &
                        "  ELSE 'Pending' " &
                        "END, " &
                        "CURDATE(), CURTIME(), " &
                        "CASE " &
                        "  WHEN Status = 'WOP' THEN Remarks " &
                        "  WHEN Status = 'Pending' OR Status IS NULL OR Status = '' THEN " &
                        "    CASE " &
                        "      WHEN ('" & activeCycle.Replace("'", "''") & "' LIKE '%Yellow%' AND Color = 'Yellow') OR ('" & activeCycle.Replace("'", "''") & "' LIKE '%Green%' AND Color = 'Green') THEN 'Not Issued' " &
                        "      ELSE 'Not Recieved' " &
                        "    END " &
                        "  ELSE '' " &
                        "END, @user " &
                        "FROM interchangeability " &
                        $"WHERE CycleName = '{activeCycle.Replace("'", "''")}' AND Status NOT IN ('Write off', 'Temp issuance') " &
                        "AND ControlNo IS NOT NULL AND ControlNo != '' " &
                        "ON DUPLICATE KEY UPDATE Status = VALUES(Status), Remarks = VALUES(Remarks), ActionBy = VALUES(ActionBy)"

            Dim pUser As New MySql.Data.MySqlClient.MySqlParameter("@user", If(currentUser, "System"))
            If mysql.ExecuteNonQuery(query, pUser) Then
                Try
                    Dim deptListQuery = "INSERT INTO department_list (Department, InstrumentName, SizeandRange, `Control No`, Color, Status, Remarks, ImportBatch, CycleName) " &
                                       $"SELECT d.Department, d.InstrumentName, d.SizeandRange, d.`Control No`, d.Color, " &
                                       "'Pending', d.Remarks, 'Cycle Rollover', " &
                                       $"'{nextCycle.Replace("'", "''")}' " &
                                       "FROM department_list d " &
                                       $"WHERE d.CycleName = '{activeCycle.Replace("'", "''")}' " &
                                       "AND d.`Control No` NOT IN (" &
                                       $"  SELECT ControlNo FROM interchangeability WHERE CycleName = '{activeCycle.Replace("'", "''")}' AND Status = 'Write off'" &
                                       ") " &
                                       "AND d.`Control No` IS NOT NULL AND d.`Control No` != '' " &
                                       "AND d.`Control No` NOT IN (" &
                                       $"  SELECT `Control No` FROM department_list WHERE CycleName = '{nextCycle.Replace("'", "''")}'" &
                                       ")"
                    mysql.ExecuteNonQuery(deptListQuery)
                Catch exDept As Exception
                    Console.WriteLine("Department List Rollover Error: " & exDept.Message)
                End Try

                mysql.SetConfigValue("ActiveCycle", nextCycle)
                mysql.CleanupStaleRFIDs(nextCycle)
                Return True
            End If
            Return False
        Catch ex As Exception
            Console.WriteLine("Rollover Error: " & ex.Message)
            Return False
        End Try
    End Function

    Public Shared Sub CheckAndRunAutomatedRollover(mysql As MySQLClass, currentUser As String)
        Try
            Dim activeCycle = mysql.GetConfigValue("ActiveCycle")
            Dim todayNaturalCycle = GetNaturalCycleNameForDate(DateTime.Today, mysql)

            If String.IsNullOrEmpty(activeCycle) Then
                mysql.SetConfigValue("ActiveCycle", todayNaturalCycle)
                Return
            End If

            If activeCycle <> todayNaturalCycle Then
                ' Perform rollover. Wait to not block UI? 
                ' This could take a moment, but it's safe to run synchronously on load or in a task.
                PerformRollover(activeCycle, todayNaturalCycle, mysql, If(currentUser, "System"))
            End If
        Catch ex As Exception
            Console.WriteLine("Auto Rollover Error: " & ex.Message)
        End Try
    End Sub

    Public Shared Sub CheckForUpcomingRolloverWarning(mysql As MySQLClass)
        Try
            Dim todayNaturalCycle = GetNaturalCycleNameForDate(DateTime.Today, mysql)
            Dim tomorrowNaturalCycle = GetNaturalCycleNameForDate(DateTime.Today.AddDays(1), mysql)

            If todayNaturalCycle <> tomorrowNaturalCycle Then
                MessageBox.Show($"Notice: The automated cycle change to {tomorrowNaturalCycle} will occur tomorrow. Please ensure all records are updated today.", "Upcoming Cycle Change", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        Catch ex As Exception
            Console.WriteLine("Warning Check Error: " & ex.Message)
        End Try
    End Sub

End Class
