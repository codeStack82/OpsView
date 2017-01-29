Imports System.Data.SqlClient 'Import SQL Capabilities
Imports System.Text

Module Module_Data
    'Module Database connections
    Private Const wv_conn = "Data Source=OKCSQLPRD0794\INST4;Initial Catalog=wv90_Reporting;Integrated Security=SSPI;"
    Private Const chkshot_conn = "Data Source=OKCSQLPRD0672\INST2;Initial Catalog=EpDmi_DrillingOA;Integrated Security=SSPI;"
    Private Const vumax_conn = "Data Source=OKCSQLPRD1021;Initial Catalog=VuMaxDR;User Id=Vmx_vumaxadmin;Password=vumaxadmin;"

    'SQL connection string
    Dim sqlConn As SqlConnection
    'Function - Get WV Active wells by idwell -returns ArrayList
    Function qry_getWVActiveWells() As Hashtable
        Dim str_Query As String

        Dim dist_wellName_Hash As New Hashtable

        str_Query = "Select DISTINCT wh.district, wh.wellName, wh.wellidc, wh.idwell " &
                    "From [dbo].[wvt_wvWellHeader] wh inner Join [dbo].[wvt_wvJobRig] wvJr on wh.idwell=wvJr.idwell " &
                        "Outer Apply(Select Top 1 jobRpt.dttmend, jobRpt.dttmstart" &
                       " From [dbo].[wvJobReport] jobRpt " &
                        "Where jobRpt.idwell = wh.idwell Order By dttmend desc) Job  " &
                    "Where Job.dttmstart > '2011-01-01 11:00:00.000' AND  syssecuritytyp <> 'NON-OP' AND wvjr.typ1 <> 'Spud' AND currentwellstatus1 <> 'PRODUCING' AND currentwellstatus1 <> 'COMPLETING'" &
                        "And currentwellstatus1 <> 'COMPLETION WOPD' AND  currentwellstatus1 <> 'COMPLETION WOPL' AND currentwellstatus1 <> 'W/O COMPLETION' AND currentwellstatus1 <> 'SHUTIN' " &
                        "And currentwellstatus1 <> 'P & A' AND currentwellstatus1 <> 'DRLD&ABAND' AND currentwellstatus1 <> 'WORKOVER' AND currentwellstatus1 <> 'SERVICE'AND currentwellstatus1 <> 'W/O PIPELINE'" &
                        "And currentwellstatus1 <> 'T/A'AND currentwellstatus1 <> 'SHUTIN WOPL'AND DateDiff(hh,GETDATE(),Job.dttmend) > -28 Order by wh.district, wh.wellName"
        sqlConn = New SqlConnection(wv_conn)

        Using (sqlConn)

            Dim sqlComm As SqlCommand = New SqlCommand(str_Query, sqlConn)
            sqlConn.Open()
            Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()

            If sqlReader.HasRows Then
                While (sqlReader.Read())
                    dist_wellName_Hash.Add(sqlReader.GetString(1), sqlReader.GetString(0))

                End While
            End If

            sqlReader.Close()
            Return dist_wellName_Hash
        End Using

    End Function

    'Function - Get WV Historic wells by idwell -returns ArrayList 
    Function qry_getWVHistoricWells() As ArrayList
        Dim str_Query As String

        Dim str_idWell As New ArrayList

        str_Query = "Select wh.wellName,  wh.idwell From [dbo].[wvt_wvWellHeader] wh where dttmspud > '2012-01-01 00:00:00.000'"
        sqlConn = New SqlConnection(wv_conn)

        Using (sqlConn)
            Dim sqlComm As SqlCommand = New SqlCommand(str_Query, sqlConn)
            sqlConn.Open()
            Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()

            If sqlReader.HasRows Then
                While (sqlReader.Read())
                    str_idWell.Add(sqlReader.GetString(1))
                End While
            End If
            sqlReader.Close()
        End Using
        Return str_idWell
    End Function

    'Function - Get WV Districts -returns ArrayList 
    Function qry_getDistrictNames() As ArrayList
        Dim str_Query As String

        Dim str_district As New ArrayList

        str_Query = "Select Distinct wh.district From [dbo].[wvt_wvWellHeader] wh where dttmspud > '2016-01-01 00:00:00.000'"
        sqlConn = New SqlConnection(wv_conn)

        Using (sqlConn)
            Dim sqlComm As SqlCommand = New SqlCommand(str_Query, sqlConn)
            sqlConn.Open()
            Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()
            If sqlReader.HasRows Then
                While (sqlReader.Read())
                    str_district.Add(sqlReader.GetString(0))
                End While
            End If
            sqlReader.Close()
        End Using
        Return str_district
    End Function

    'Function - Get ChkShot well by PN -returns ArrayList ~~~~TODO: Finish
    Function qry_getChkShotWell() As ArrayList 'ByVal PN As Integer
        Dim str_Query As String

        Dim str_idWell As New ArrayList

        str_Query = "SELECT SVY.DrillingStatus " &
                      "  , SVY.PtbMeasuredDepth, SVY.MeasuredDepth " &
                      "  , SVY.PtbInclination " &
                      "  , SVY.PtbAzimuth " &
                      "  , SVY.AboveBelowPlan " &
                      "  , SVY.LeftRightPlan " &
                      "  , SVY.DogLegSeverity " &
                      " , SVY.DogLegNeeded " &
                      "  , SVY.WellHeaderID " &
                      " From chkshot.vw_CurrentWell AS CW with (nolock) " &
                      " OUTER APPLY(" &
                      "      SELECT TOP 1 DrillingStatus, PtbMeasuredDepth, MeasuredDepth " &
                      "                  , PtbInclination" &
                      "                  , PtbAzimuth " &
                      "                  , AboveBelowPlan , LeftRightPlan " &
                      "                  , DogLegSeverity " &
                      "                  , DogLegNeeded " &
                      "                 , WellHeaderID " &
                      "      From chkshot.vw_CHKShot_Moblize_Active_Surveys AS SVY with (nolock)" &
                      "      Where SVY.PropertyNumber = CW.PropertyNumber" &
                      "      Order By SVY.MeasuredDepth DESC " &
                      "  ) AS SVY " &
                      "  WHERE CW.PropertyNumber = '657699'"


        sqlConn = New SqlConnection(chkshot_conn)

        Using (sqlConn)
            Dim sqlComm As SqlCommand = New SqlCommand(str_Query, sqlConn)
            sqlConn.Open()
            Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()

            If sqlReader.HasRows Then
                While (sqlReader.Read())
                    'Will need to build output string for Survey

                    MsgBox(CType(sqlReader.GetDouble(1), String))

                    'str_idWell.Add(sqlReader.GetString(1))
                End While
            End If
            sqlReader.Close()
        End Using
        Return str_idWell
    End Function

    'Function - Get VuMax well data by idwell -returns ArrayList 
    Function qry_getVumaxWellData() As ArrayList
        Dim str_Query As String

        Dim str_idWell As New ArrayList

        str_Query = "Select DATA_INDEX from timeLog657699Wellbore1TimeLog1"
        sqlConn = New SqlConnection(vumax_conn)

        Using (sqlConn)
            Dim sqlComm As SqlCommand = New SqlCommand(str_Query, sqlConn)
            sqlConn.Open()
            Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()

            If sqlReader.HasRows Then
                While (sqlReader.Read())

                    MsgBox(sqlReader.GetDecimal(0))
                    'str_idWell.Add(sqlReader.GetString(0))
                End While
            End If
            sqlReader.Close()
        End Using
        Return str_idWell
    End Function

    'Function - District linq query
    Function linq_getDistrictNames() As Array
        Dim db As New LINQ_WellHeaderDataContext()

        Dim Districts = From d In db.wvt_wvWellHeaders
                        Where d.dttmspud > Today.AddMonths(-3) And d.district IsNot "DEVONIAN" And d.district IsNot "WILLISTON"
                        Select d.district Distinct

        Return Districts.ToArray()
    End Function




End Module
