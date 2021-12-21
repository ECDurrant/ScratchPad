
Imports System.Data.OleDb

Imports System.Threading







Public Class ScratchPadMain


    ''Store a Call
    Dim SQL As String
    Dim con As New OleDbConnection

    ''RefNum
    Dim SQL1 As String
    Dim con1 As New OleDbConnection
    ''Call Counter
    Dim CallCounter As Integer
    ''LogLater Counter
    Dim LogLaterCouner As Integer
    ''TicketLogged Counter
    Dim TicketsLogged As Integer
    ''Confluence Launcher
    Dim conflu As String = "https://confluence.fadv.com/"
    ''Workday Launcher
    Dim workd As String = "https://wd5.myworkday.com/fadv/fx/home.flex"

    ''Store Call Thread
    Dim TemplateStoreThread As System.Threading.Thread

    ''Store Call Thread
    Dim StoreCallThread As System.Threading.Thread
    ''Log later Store Call
    Dim LoglaterStoreCall As System.Threading.Thread
    ''Call List
    Dim CallList_Rand As System.Threading.Thread
    ''Call Count 
    Dim Call_C As System.Threading.Thread

    ''TestThread
    Dim RefNumT As System.Threading.Thread

    Public Sub Call_Counter()

        Try

            ''Call Counter code

            CallCounter = CallCounter + 1

            lblcCounter.Text = CallCounter.ToString

        Catch ex As Exception


            MsgBox(ex.Message)

        End Try

    End Sub


    Public Sub Log_Ticket_Counter()



        TicketsLogged = TicketsLogged + 1

        lblTicketsLoggedNumber.Text = TicketsLogged.ToString








    End Sub


    Public Sub newref()












    End Sub

    
    Public Sub Reference()

        Try

        
        Static random000 As New Random

        Dim Refresult As Integer

            Refresult = random000.Next(7000, 30000)

        If lblAgentName.Text = "Eric Durrant" Then

                lblSPCRefNum.Text = Refresult + 111

        ElseIf lblAgentName.Text = "Lucio Benzor" Then

                lblSPCRefNum.Text = Refresult + 222

        ElseIf lblAgentName.Text = "Vernett Hines" Then

                lblSPCRefNum.Text = Refresult + 333

        ElseIf lblAgentName.Text = "Brittany Brady" Then

                lblSPCRefNum.Text = Refresult + 444

        ElseIf lblAgentName.Text = "Lauren Lee" Then

                lblSPCRefNum.Text = Refresult + 555

        ElseIf lblAgentName.Text = "Adrian Davis" Then


                lblSPCRefNum.Text = Refresult + 666

        ElseIf lblAgentName.Text = "Kimberle Lawrence" Then

                lblSPCRefNum.Text = Refresult + 777

        Else

                lblSPCRefNum.Text = Refresult + 8888

        End If


        Catch ex As Exception

            MsgBox(ex.Message)


        End Try



    End Sub







    Public Sub LogStoreCall()




        Reference()

        Try

            con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\SPC\ScratchPad Database.accdb")



            con.Open()



            Dim SQL As String = "INSERT INTO [July - December 2015] ([Platform], [AgentName], [JobTitle],[FirstName], [LastName], [Phone], [Email],[OEmail],[UserID],[ClientID],[OrderID],[ApplicationID],[CallDetail],[DateOfCall],[TimeOfCall],[Reference]) Values ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

            Using cmd As New OleDbCommand(SQL, con)



                ''Selecting the Platform

                If radNewHire.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "CVS Guardian")

                ElseIf radManager.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "CVS Guardian")

                ElseIf radContractor.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "CVS Guardian")

                End If

                If radClient.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "Enterprise Advantage")

                ElseIf radApplicantEnt.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "Enterprise Advantage")

                ElseIf radOther.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "Transfer")

                ElseIf radFingerPrinting.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "FingerPrinting")

                ElseIf radPROM.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "PROM")

                End If

                ''==================================================================================================
                cmd.Parameters.AddWithValue("@p2", lblAgentName.Text)

                ''Select Job Title

                If radNewHire.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "New Hire/CVS")

                ElseIf radManager.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "Store Manager/CVS")


                ElseIf radContractor.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "Field colleague trainer/CVS")

                ElseIf radFingerPrinting.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "FingerPrinting")

                ElseIf radOther.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "Transfer")


                ElseIf radClient.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "Enterprise Client")


                ElseIf radApplicantEnt.Checked = True Then


                    cmd.Parameters.AddWithValue("@p3", "Enterprise Applicant")

                ElseIf radPROM.Checked = True Then


                    cmd.Parameters.AddWithValue("@p3", "PROM")




                End If

                ''=============================================================================================

                cmd.Parameters.AddWithValue("@p4", txtFirstName.Text)
                cmd.Parameters.AddWithValue("@p5", txtLastName.Text)
                cmd.Parameters.AddWithValue("@p6", txtPhone.Text)




                ''Decide what textbox the Email will come from

                If radNewHire.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)

                ElseIf radManager.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtNewHireEmail.Text)


                ElseIf radContractor.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtNewHireEmail.Text)

                ElseIf radFingerPrinting.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)

                ElseIf radOther.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)


                ElseIf radClient.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)


                ElseIf radApplicantEnt.Checked = True Then


                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)

                ElseIf radPROM.Checked = True Then


                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)


                End If


                ''=============================================================================================


                cmd.Parameters.AddWithValue("@p8", "N/a")
                cmd.Parameters.AddWithValue("@p9", txtUserID.Text)
                cmd.Parameters.AddWithValue("@p10", txtClientID.Text)
                cmd.Parameters.AddWithValue("@p11", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@p12", txtApplicationID.Text)
                cmd.Parameters.AddWithValue("@p13", txtCallDetail.Text)
                cmd.Parameters.AddWithValue("@p14", lblNewDate.Text)
                cmd.Parameters.AddWithValue("@p15", lblNewTime.Text)
                cmd.Parameters.AddWithValue("@p16", lblSPCRefNum.Text)


                cmd.ExecuteNonQuery()

                con.Close()



            End Using


        Catch ex As OleDbException

            MsgBox(ex.Message)

        Catch ex As SystemException

            MsgBox(ex.Message)

        Catch ex As Exception

            MsgBox(ex.Message)

            MsgBox("back file 2 global")

        End Try





    End Sub





    Public Sub CallList()

        Try


            Reference()


            ''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            ListBox1.Items.Add(lblNewDate.Text + " / " + lblNewTime.Text)
            ListBox1.Items.Add("Call#: " + lblcCounter.Text)
            ListBox1.Items.Add("SPC Reference #: " + lblSPCRefNum.Text)

            If radNewHire.Checked = True Then

                ListBox1.Items.Add("CVS New Hire")

            ElseIf radManager.Checked = True Then


                ListBox1.Items.Add("CVS Store Manager")

            ElseIf radContractor.Checked = True Then

                ListBox1.Items.Add("Contractor Advantage")

            ElseIf radOther.Checked = True Then

                ListBox1.Items.Add("Applicant Transfer")



            ElseIf radFingerPrinting.Checked = True Then

                ListBox1.Items.Add("Fingerprinting")

            ElseIf radClient.Checked = True Then

                ListBox1.Items.Add("Enterprise Client")


            ElseIf radApplicantEnt.Checked = True Then

                ListBox1.Items.Add("Enterprise Applicant")


            ElseIf radPROM.Checked = True Then

                ListBox1.Items.Add("PROM")
                ' 1-1138742798839



            End If




            ListBox1.Items.Add(txtFirstName.Text)
            ListBox1.Items.Add(txtLastName.Text)
            ListBox1.Items.Add(txtPhone.Text)

            If radNewHire.Checked = True Then

                ListBox1.Items.Add(txtEmail.Text)
                ListBox1.Items.Add("1-1134812539086")

            ElseIf radManager.Checked = True Then

                ListBox1.Items.Add(txtUserID.Text)
                ListBox1.Items.Add(txtNewHireEmail.Text)

                If txtAppName.Text = "" Then

                    Log_Later.ListBox1.Items.Add("N/a")

                Else
                    Log_Later.ListBox1.Items.Add(txtAppName.Text)

                End If

                ListBox1.Items.Add("1-1134812539086")

            ElseIf radContractor.Checked = True Then

                ListBox1.Items.Add(txtEmail.Text)
                ListBox1.Items.Add(txtUserID.Text)
                ListBox1.Items.Add(txtOrderID.Text)
                ListBox1.Items.Add(txtAccountName.Text)
                If txtAppName.Text = "" Then

                    Log_Later.ListBox1.Items.Add("N/a")

                Else
                    Log_Later.ListBox1.Items.Add(txtAppName.Text)

                End If
         
                ''Transfer ===============================================================================================================================================================================
            ElseIf radOther.Checked = True Then

                If txtUserID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtUserID.Text)

                End If

                If txtClientID.Text = "" Then

                    ListBox1.Items.Add("N/a")
                Else

                    ListBox1.Items.Add(txtClientID.Text)

                End If

                If txtAccountName.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtAccountName.Text)

                End If

                If txtEmail.Text = "" Then

                    ListBox1.Items.Add("N/a")
                Else

                    ListBox1.Items.Add(txtEmail.Text)

                End If



                ''Fingerprinting =========================================================================================================================================================


            ElseIf radFingerPrinting.Checked = True Then

                If txtEmail.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else

                    ListBox1.Items.Add(txtEmail.Text)

                End If



                If txtAccountName.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else

                    ListBox1.Items.Add(txtAccountName.Text)

                End If




                ''Enterprise Client =====================================================================================================================================================

            ElseIf radClient.Checked = True Then

                If txtUserID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtUserID.Text)

                End If


                If txtClientID.Text = "" Then

                    ListBox1.Items.Add("N/a")
                Else

                    ListBox1.Items.Add(txtClientID.Text)

                End If

                If txtOrderID.Text = "" Then


                    ListBox1.Items.Add("N/a")

                Else

                    ListBox1.Items.Add(txtOrderID.Text)

                End If


                If txtEmail.Text = "" Then

                    ListBox1.Items.Add("N/a")
                Else

                    ListBox1.Items.Add(txtEmail.Text)

                End If

                If txtAppName.Text = "" Then

                    Log_Later.ListBox1.Items.Add("N/a")

                Else
                    Log_Later.ListBox1.Items.Add(txtAppName.Text)

                End If

                ListBox1.Items.Add(txtAccountName.Text)

                '' Enterprise Applicant ==============================================================================================================================================

            ElseIf radApplicantEnt.Checked = True Then

                If txtApplicationID.Text = "" Then

                    ListBox1.Items.Add("N/a")
                Else

                    ListBox1.Items.Add(txtApplicationID.Text)

                End If

                If txtOrderID.Text = "" Then

                    ListBox1.Items.Add("N/a")
                Else

                    ListBox1.Items.Add(txtOrderID.Text)

                End If

                If txtEmail.Text = "" Then

                    ListBox1.Items.Add("N/a")
                Else

                    ListBox1.Items.Add(txtEmail.Text)

                End If

                ListBox1.Items.Add(txtAccountName.Text)

                '===========PROM==============================================================================================================================================

            ElseIf radPROM.Checked = True Then

                If txtUserID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else

                    ListBox1.Items.Add(txtUserID.Text)

                End If

                If txtEmail.Text = "" Then

                    ListBox1.Items.Add("N/a")


                Else

                    ListBox1.Items.Add(txtEmail.Text)

                End If


                ListBox1.Items.Add("1-1138742798839")




            End If

            ListBox1.Items.Add(txtCallDetail.Text)
            ListBox1.Items.Add("-----------------------------------------------------------------------------------------------------------------------------")
            ListBox1.Items.Add("")


        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try


    End Sub



    Public Sub Backend()

        txtFirstName.Text = BackEnder.txtName.Text

        txtLastName.Text = BackEnder.txtLast.Text

        txtPhone.Text = BackEnder.txtPhone.Text

        txtEmail.Text = BackEnder.txtEmail.Text







    End Sub



    Private Sub StoreCall()


        Try





            con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\SPC\ScratchPad Database.accdb")



            con.Open()



            Dim SQL As String = "INSERT INTO [July - December 2015] ([Platform], [AgentName], [JobTitle],[FirstName], [LastName], [Phone], [Email],[OEmail],[UserID],[ClientID],[OrderID],[ApplicationID],[CallDetail],[DateOfCall],[TimeOfCall],[Reference]) Values ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?)"

            Using cmd As New OleDbCommand(SQL, con)






                ''Selecting the Platform

                If radNewHire.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "CVS Guardian")

                ElseIf radManager.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "CVS Guardian")

                ElseIf radContractor.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "CVS Guardian")

                End If

                If radClient.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "Enterprise Advantage")

                ElseIf radApplicantEnt.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "Enterprise Advantage")

                ElseIf radOther.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "Transfer")

                ElseIf radFingerPrinting.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "FingerPrinting")

                ElseIf radPROM.Checked = True Then

                    cmd.Parameters.AddWithValue("@p1", "PROM")

                End If

                ''==================================================================================================
                cmd.Parameters.AddWithValue("@p2", lblAgentName.Text)

                ''Select Job Title

                If radNewHire.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "New Hire/CVS")

                ElseIf radManager.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "Store Manager/CVS")


                ElseIf radContractor.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "Field colleague trainer/CVS")

                ElseIf radFingerPrinting.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "FingerPrinting")

                ElseIf radOther.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "Transfer")


                ElseIf radClient.Checked = True Then

                    cmd.Parameters.AddWithValue("@p3", "Enterprise Client")


                ElseIf radApplicantEnt.Checked = True Then


                    cmd.Parameters.AddWithValue("@p3", "Enterprise Applicant")

                ElseIf radPROM.Checked = True Then


                    cmd.Parameters.AddWithValue("@p3", "PROM")


                End If

                ''=============================================================================================

                cmd.Parameters.AddWithValue("@p4", txtFirstName.Text)
                cmd.Parameters.AddWithValue("@p5", txtLastName.Text)
                cmd.Parameters.AddWithValue("@p6", txtPhone.Text)




                ''Decide what textbox the Email will come from

                If radNewHire.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)

                ElseIf radManager.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtNewHireEmail.Text)


                ElseIf radContractor.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtNewHireEmail.Text)

                ElseIf radFingerPrinting.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)

                ElseIf radOther.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)


                ElseIf radClient.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)


                ElseIf radApplicantEnt.Checked = True Then


                    cmd.Parameters.AddWithValue("@p7", txtEmail.Text)


                ElseIf radPROM.Checked = True Then

                    cmd.Parameters.AddWithValue("@p7", txtNewHireEmail.Text)


                End If


                ''=============================================================================================


                cmd.Parameters.AddWithValue("@p7", "N/a")
                cmd.Parameters.AddWithValue("@p9", txtUserID.Text)
                cmd.Parameters.AddWithValue("@p10", txtClientID.Text)
                cmd.Parameters.AddWithValue("@p11", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@p12", txtApplicationID.Text)
                cmd.Parameters.AddWithValue("@p13", txtCallDetail.Text)
                cmd.Parameters.AddWithValue("@p14", lblNewDate.Text)
                cmd.Parameters.AddWithValue("@p15", lblNewTime.Text)
                cmd.Parameters.AddWithValue("@p16", lblSPCRefNum.Text)


                cmd.ExecuteNonQuery()

                con.Close()



            End Using


            ''Error Checking
        Catch ex As OleDbException
            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()


                MsgBox("The connection to the P drive was interupted..@ store call procedure")


            End If

        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            MsgBox("system error at store call")


            MsgBox("Attempted to fix connection and store information")

        Catch ex As Exception


            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try




    End Sub

    Public Sub LogLater()

        Try

            ''Store call information to Call List



            ListBox1.Items.Add(lblNewDate.Text + " / " + lblNewTime.Text)
            ListBox1.Items.Add("Call#: " + lblcCounter.Text)
            ListBox1.Items.Add("SPC Reference #: " + lblSPCRefNum.Text)

            If radNewHire.Checked = True Then

                ListBox1.Items.Add("CVS New Hire")

            ElseIf radManager.Checked = True Then


                ListBox1.Items.Add("CVS Store Manager")

            ElseIf radContractor.Checked = True Then

                ListBox1.Items.Add("Field Colleague Trainer")

            ElseIf radOther.Checked = True Then

                ListBox1.Items.Add("Warm Transfer")


            ElseIf radFingerPrinting.Checked = True Then

                ListBox1.Items.Add("FingerPrinting")


            ElseIf radClient.Checked = True Then

                ListBox1.Items.Add("Enterprise Client")


            ElseIf radApplicantEnt.Checked = True Then

                ListBox1.Items.Add("Enterprise Applicant")

            ElseIf radPROM.Checked = True Then

                ListBox1.Items.Add("PROM")

            End If



            ListBox1.Items.Add(txtFirstName.Text)
            ListBox1.Items.Add(txtLastName.Text)
            ListBox1.Items.Add(txtPhone.Text)

            If radNewHire.Checked = True Then

                ListBox1.Items.Add(txtEmail.Text)
                ListBox1.Items.Add("1-1134812539086")

            ElseIf radManager.Checked = True Then

                ListBox1.Items.Add(txtUserID.Text)
                ListBox1.Items.Add(txtNewHireEmail.Text)
                ListBox1.Items.Add("1-1134812539086")

            ElseIf radContractor.Checked = True Then

                ListBox1.Items.Add(txtUserID.Text)
                ListBox1.Items.Add(txtNewHireEmail.Text)
                ListBox1.Items.Add("1-1134812539086")





                '' 
            ElseIf radOther.Checked = True Then

                If txtUserID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtUserID.Text)

                End If

                If txtClientID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtClientID.Text)

                End If

                If txtOtherOptionTxt.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtOtherOptionTxt.Text)

                End If


                ListBox1.Items.Add(txtAccountName.Text)


            ElseIf radFingerPrinting.Checked = True Then

                ListBox1.Items.Add(txtEmail.Text)
                ListBox1.Items.Add(txtAccountName.Text)


                '' Enterprise Client =======================================================================================================================

            ElseIf radClient.Checked = True Then

                If txtUserID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtUserID.Text)

                End If


                If txtClientID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtClientID.Text)

                End If


                If txtOtherOptionTxt.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtOtherOptionTxt.Text)

                End If


                If txtOrderID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtOrderID.Text)

                End If

                ListBox1.Items.Add(txtAccountName.Text)


                ''Enterprise Applicant ===================================================================================================================================

            ElseIf radApplicantEnt.Checked = True Then

                If txtEmail.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtEmail.Text)

                End If


                If txtApplicationID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtApplicationID.Text)

                End If


                If txtOrderID.Text = "" Then

                    ListBox1.Items.Add("N/a")

                Else
                    ListBox1.Items.Add(txtOrderID.Text)

                End If


                ListBox1.Items.Add(txtAccountName.Text)

            End If







            ListBox1.Items.Add(txtCallDetail.Text)
            ListBox1.Items.Add("")
            ListBox1.Items.Add("!!!!!!!!!!!!!!--- Log This Call--!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            ListBox1.Items.Add("")
            ListBox1.Items.Add("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
            ListBox1.Items.Add("")



            ''Error Checking
        Catch ex As OleDbException

            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()



                MsgBox("The connection to the P drive was interupted..@ log later procedure")


            End If

        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            MsgBox("system error at log later prodecure")


        Catch ex As Exception


            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub

    Public Sub FillCombo()

        Try



            sqltemp2 = "SELECT * FROM [Personal Templates] WHERE User='" & lblAgentName.Text & "' "

            Dim cmdtemp As New OleDb.OleDbCommand

            '  cmdtemp.CommandText = sqltemp
            cmdtemp.CommandText = sqltemp2
            cmdtemp.Connection = contemp


            readertemp = cmdtemp.ExecuteReader

            While (readertemp.Read())

                NEWBOX.Items.Add(readertemp("TemplateName"))



            End While




            cmdtemp.Dispose()
            readertemp.Close()


            ''Error Checking
        Catch ex As OleDbException

            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()


                MsgBox("The connection to the P drive was interupted..@ fill combo procedure")


            End If

        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            '  MsgBox("system error at fill combo procedure")


        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub

    Public Sub FillCombo2()

        Try

            ' sqltemp = "SELECT * FROM Templates"

            ' sqltemp = "SELECT * FROM Templates WHERE TemplateName='" & NEWBOX.Text & " ' "

            ' sqltemp3 = "SELECT * FROM Templates WHERE UserP= ' " & Label9.Text & "' "

            sqltemp3 = "SELECT * FROM  [CVS Templates]"


            Dim cmdtemp3 As New OleDb.OleDbCommand

            '  cmdtemp.CommandText = sqltemp
            cmdtemp3.CommandText = sqltemp3
            cmdtemp3.Connection = contemp3


            readertemp3 = cmdtemp3.ExecuteReader

            While (readertemp3.Read())

                ComboBox2.Items.Add(readertemp3("TemplateName"))



            End While

            cmdtemp3.Dispose()
            readertemp3.Close()


            ''Error Checking
        Catch ex As OleDbException

          


        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            '   MsgBox("system error at fill combo procedure 2")




        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try

    End Sub





    Public Sub ScratchPadMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

     
        'Load templates
        TemplateMod.connecttemp()
        FillCombo()

        TemplateMod.connecttemp3()
        FillCombo2()

        ''Default Focus
        Me.ActiveControl = txtFirstName

        ''Threading Error Check
        Control.CheckForIllegalCrossThreadCalls = False

        ''Running Time
        Time.Enabled = True

        ''Center to Screen
        Me.CenterToScreen()

        ''Open Siebel
        Siebel.Show()

            '   BackEnder.Show()

        '  StartTimerForCombo.Enabled = True

            Timer1.Enabled = True


            ''Error Checking





        Catch ex As OleDbException

            MsgBox(ex.Message)

   
        Catch ex As SystemException

            MsgBox(ex.Message)


        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")


       
        End Try



    End Sub

    Public Sub SendtoBackender()





        BackEnder.txtName.Text = txtFirstName.Text

        BackEnder.txtLast.Text = txtLastName.Text

        BackEnder.txtPhone.Text = txtPhone.Text

        BackEnder.txtEmail1.Text = txtEmail.Text

        BackEnder.txtEmail3.Text = txtNewHireEmail.Text

        BackEnder.txtEmpID.Text = txtUserID.Text

        BackEnder.txtappname.Text = txtAppName.Text



      

        BackEnder.txtClientId.Text = txtClientID.Text

        BackEnder.txtOrderID.Text = txtOrderID.Text

        BackEnder.txtAccountName.Text = txtAccountName.Text

        BackEnder.txtAppID.Text = txtApplicationID.Text


        BackEnder.txtCallDetail.Text = txtCallDetail.Text




    End Sub


    Public Sub ClearBackender()



        BackEnder.txtName.Clear()


        BackEnder.txtName.Clear()


        BackEnder.txtLast.Clear()


        BackEnder.txtPhone.Clear()


        BackEnder.txtEmail.Clear()


        BackEnder.txtEmail1.Clear()


        BackEnder.txtEmail3.Clear()


        BackEnder.txtEmpID.Clear()





        BackEnder.txtVendorID.Clear()


        BackEnder.txtClientDetail.Clear()

        BackEnder.txtJobTitle.Clear()


        BackEnder.txtClientId.Clear()

        BackEnder.txtOrderID.Clear()


        BackEnder.txtAccountName.Clear()


        BackEnder.txtAppID.Clear()



        BackEnder.txtCallDetail.Clear()




    End Sub



    Public Sub backuptotxt2()

        Try




            Dim Globalbackuptotxt As System.IO.StreamWriter


            Globalbackuptotxt = My.Computer.FileSystem.OpenTextFileWriter("P:\SPC\GlobalSPC.txt", True)



            ''Job title Radio buttin Code to text goes here

            Globalbackuptotxt.WriteLine("Date/Time: " & lblNewDate.Text & " - " & lblNewTime.Text)

            Globalbackuptotxt.WriteLine("Call#: " & lblcCounter.Text)

            Globalbackuptotxt.WriteLine("FADV Agent: " & lblAgentName.Text)

            If radNewHire.Checked = True Then



                Globalbackuptotxt.WriteLine("Job Title: CVS - New Hire")


            ElseIf radManager.Checked = True Then

                Globalbackuptotxt.WriteLine("Job Title: CVS - User")


            ElseIf radContractor.Checked = True Then

                Globalbackuptotxt.WriteLine("Job Title: Contractor Advantage")

            ElseIf radPROM.Checked = True Then

                Globalbackuptotxt.WriteLine("Job Title: PROM")

            ElseIf radFingerPrinting.Checked = True Then

                Globalbackuptotxt.WriteLine("Job Title: Fringerprinting")


            ElseIf radClient.Checked = True Then

                Globalbackuptotxt.WriteLine("Job Title: Enterprise Advantage Client")


            ElseIf radApplicantEnt.Checked = True Then

                Globalbackuptotxt.WriteLine("Job Title: Enterprise Advantage Applicant")


            ElseIf radOther.Checked = True Then

                Globalbackuptotxt.WriteLine("Job Title: Applicant Transfer")

            End If



            Globalbackuptotxt.WriteLine("First Name:" & txtFirstName.Text)
            Globalbackuptotxt.WriteLine("Last Name:" & txtLastName.Text)
            Globalbackuptotxt.WriteLine("Phone#:" & txtPhone.Text)


            ''Conditions

            If radNewHire.Checked = True Then

                Globalbackuptotxt.WriteLine("Email:" & txtEmail.Text)

            End If



            If radManager.Checked = True Then
                Globalbackuptotxt.WriteLine("Email:" & txtNewHireEmail.Text)
                Globalbackuptotxt.WriteLine("User ID:" & txtUserID.Text)

            End If



            If radContractor.Checked = True Then

                Globalbackuptotxt.WriteLine("Email:" & txtEmail.Text)
                Globalbackuptotxt.WriteLine("User ID:" & txtUserID.Text)
                Globalbackuptotxt.WriteLine("Order ID:" & txtOrderID.Text)
                Globalbackuptotxt.WriteLine("Account Name:" & txtAccountName.Text)


            End If

            If radPROM.Checked = True Then

                Globalbackuptotxt.WriteLine("Email:" & txtEmail.Text)
                Globalbackuptotxt.WriteLine("Vendor ID:" & txtUserID.Text)



            End If


            If radClient.Checked = True Then

                Globalbackuptotxt.WriteLine("User ID:" & txtUserID.Text)
                Globalbackuptotxt.WriteLine("Client ID:" & txtClientID.Text)
                Globalbackuptotxt.WriteLine("Order ID:" & txtOrderID.Text)
                Globalbackuptotxt.WriteLine("Account Name:" & txtAccountName.Text)
                Globalbackuptotxt.WriteLine("Email:" & txtEmail.Text)




            End If


            If radApplicantEnt.Checked = True Then

                Globalbackuptotxt.WriteLine("Email:" & txtEmail.Text)
                Globalbackuptotxt.WriteLine("App ID:" & txtApplicationID.Text)
                Globalbackuptotxt.WriteLine("Order ID:" & txtOrderID.Text)
                Globalbackuptotxt.WriteLine("Account Name:" & txtAccountName.Text)


            End If






            If radFingerPrinting.Checked Then

                Globalbackuptotxt.WriteLine("Email:" & txtEmail.Text)
                Globalbackuptotxt.WriteLine("Account Name:" & txtAccountName.Text)


            End If

            If radOther.Checked Then


                Globalbackuptotxt.WriteLine("User ID:" & txtUserID.Text)
                Globalbackuptotxt.WriteLine("Account Name:" & txtAccountName.Text)
                Globalbackuptotxt.WriteLine("Email:" & txtEmail.Text)



            End If


            Globalbackuptotxt.WriteLine("Notes:" & txtCallDetail.Text)



            Globalbackuptotxt.WriteLine("----------------------------------------------------------------------------------------------------------------------")
            Globalbackuptotxt.Close()


        Catch ex As OleDbException

            MsgBox("There as been an connection break, please restart for the drop down to load contents")

            MsgBox(ex.Message)

        Catch ex As SystemException

            MsgBox(ex.Message)

        Catch ex As Exception

            MsgBox(ex.Message)

            MsgBox("back file 2 global")
        End Try

    End Sub

    Public Sub backuptotxt()

        Try




            Dim backuptotxt As System.IO.StreamWriter


            backuptotxt = My.Computer.FileSystem.OpenTextFileWriter("P:\SPC\Global Scratchpad Backup Information.txt", True)
            backuptotxt = My.Computer.FileSystem.OpenTextFileWriter("C:\Scratchpad Backup Folder\My Backup CallList.txt", True)


            ''Job title Radio buttin Code to text goes here

            backuptotxt.WriteLine("Date/Time: " & lblNewDate.Text & " - " & lblNewTime.Text)

            backuptotxt.WriteLine("Call#: " & lblcCounter.Text)

            backuptotxt.WriteLine("FADV Agent: " & lblAgentName.Text)

            If radNewHire.Checked = True Then



                backuptotxt.WriteLine("Job Title: CVS - New Hire")


            ElseIf radManager.Checked = True Then

                backuptotxt.WriteLine("Job Title: CVS - User")


            ElseIf radContractor.Checked = True Then

                backuptotxt.WriteLine("Job Title: Contractor Advantage")

            ElseIf radPROM.Checked = True Then

                backuptotxt.WriteLine("Job Title: PROM")

            ElseIf radFingerPrinting.Checked = True Then

                backuptotxt.WriteLine("Job Title: Fringerprinting")


            ElseIf radClient.Checked = True Then

                backuptotxt.WriteLine("Job Title: Enterprise Advantage Client")


            ElseIf radApplicantEnt.Checked = True Then

                backuptotxt.WriteLine("Job Title: Enterprise Advantage Applicant")


            ElseIf radOther.Checked = True Then

                backuptotxt.WriteLine("Job Title: Applicant Transfer")

            End If



            backuptotxt.WriteLine("First Name:" & txtFirstName.Text)
            backuptotxt.WriteLine("Last Name:" & txtLastName.Text)
            backuptotxt.WriteLine("Phone#:" & txtPhone.Text)


            ''Conditions

            If radNewHire.Checked = True Then

                backuptotxt.WriteLine("Email:" & txtEmail.Text)

            End If



            If radManager.Checked = True Then
                backuptotxt.WriteLine("Email:" & txtNewHireEmail.Text)
                backuptotxt.WriteLine("User ID:" & txtUserID.Text)

            End If



            If radContractor.Checked = True Then

                backuptotxt.WriteLine("Email:" & txtEmail.Text)
                backuptotxt.WriteLine("User ID:" & txtUserID.Text)
                backuptotxt.WriteLine("Order ID:" & txtOrderID.Text)
                backuptotxt.WriteLine("Account Name:" & txtAccountName.Text)


            End If

            If radPROM.Checked = True Then

                backuptotxt.WriteLine("Email:" & txtEmail.Text)
                backuptotxt.WriteLine("Vendor ID:" & txtUserID.Text)



            End If


            If radClient.Checked = True Then

                backuptotxt.WriteLine("User ID:" & txtUserID.Text)
                backuptotxt.WriteLine("Client ID:" & txtClientID.Text)
                backuptotxt.WriteLine("Order ID:" & txtOrderID.Text)
                backuptotxt.WriteLine("Account Name:" & txtAccountName.Text)
                backuptotxt.WriteLine("Email:" & txtEmail.Text)




            End If


            If radApplicantEnt.Checked = True Then

                backuptotxt.WriteLine("Email:" & txtEmail.Text)
                backuptotxt.WriteLine("App ID:" & txtApplicationID.Text)
                backuptotxt.WriteLine("Order ID:" & txtOrderID.Text)
                backuptotxt.WriteLine("Account Name:" & txtAccountName.Text)


            End If






            If radFingerPrinting.Checked Then

                backuptotxt.WriteLine("Email:" & txtEmail.Text)
                backuptotxt.WriteLine("Account Name:" & txtAccountName.Text)


            End If

            If radOther.Checked Then


                backuptotxt.WriteLine("User ID:" & txtUserID.Text)
                backuptotxt.WriteLine("Account Name:" & txtAccountName.Text)
                backuptotxt.WriteLine("Email:" & txtEmail.Text)



            End If


            backuptotxt.WriteLine("Notes:" & txtCallDetail.Text)



            backuptotxt.WriteLine("----------------------------------------------------------------------------------------------------------------------")
            backuptotxt.Close()


        Catch ex As OleDbException



            MsgBox(ex.Message)

            MsgBox("back file 2 global")
        Catch ex As SystemException

            MsgBox(ex.Message)

            MsgBox("back file 2 global")
        Catch ex As Exception

            MsgBox(ex.Message)

            MsgBox("back file 2 global")

        End Try



    End Sub




    Private Sub btnStore_Click(sender As Object, e As EventArgs) Handles btnStore.Click


        Try



            Me.Cursor = Cursors.WaitCursor

            '' Job title must be selected

            If radManager.Checked = False And radNewHire.Checked = False And radContractor.Checked = False And radOther.Checked = False And radClient.Checked = False And radApplicantEnt.Checked = False And radFingerPrinting.Checked = False And radPROM.Checked = False Then

                MessageBox.Show("Please be advised that a Job title must be selected in order to save this call, please select a Job Title and try again", "Warning", MessageBoxButtons.RetryCancel)

                Me.Cursor = Cursors.Hand


            Else


                '' Required fields must be filled out


                If txtFirstName.Text = "" Then

                    MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                    Me.ActiveControl = txtFirstName

                    Me.Cursor = Cursors.Hand

                Else


                    '' Required fields must be filled out

                    If txtLastName.Text = "" Then

                        MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                        Me.ActiveControl = txtLastName

                        Me.Cursor = Cursors.Hand

                    Else

                        '' Required fields must be filled out

                        If txtPhone.Text = "" Then


                            MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                            Me.ActiveControl = txtPhone

                            Me.Cursor = Cursors.Hand

                        Else



                            ''No Dulicate Calls

                            If lblSNoti.Visible = True Then

                                MessageBox.Show("Be Advised this Call was stored already; In order to reduce duplicate entries in the database you must select 'New Call' before saving ", "Warning", MessageBoxButtons.RetryCancel)


                                Me.Cursor = Cursors.Hand

                            Else


                                ''Backup
                                backuptotxt()

                                backuptotxt2()


                                ''Call List and Reference Code

                                CallList_Rand = New System.Threading.Thread(AddressOf CallList)

                                CallList_Rand.Start()



                                Call_C = New System.Threading.Thread(AddressOf Call_Counter)

                                Call_C.Start()



                                ''Call List Timer
                                CallListTimer.Enabled = True





                                ''Store Call Thread 

                                StoreCallThread = New System.Threading.Thread(AddressOf StoreCall)

                                StoreCallThread.Start()




                                ''============Changing the labels ==================================================================


                  


                                'Ref label

                                lblReflabel.Visible = True
                                lblSPCRefNum.Visible = True

                                ''Stored Notification

                                lblSNoti.Visible = True


                                ''Text to button code


                                lblDailyCallC.Visible = True
                                lblcCounter.Visible = True

                                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue
                                Me.lblAppName.ForeColor = System.Drawing.Color.Blue

                                lblOrderID.ForeColor = Color.Blue
                                lblAppID.ForeColor = Color.Blue
                                lblAccountName.ForeColor = Color.Blue



                                Me.Cursor = Cursors.Hand





                            End If
                        End If
                    End If
                End If
            End If



        Catch ex As SyntaxErrorException


        Catch ex As SystemException

            MsgBox("system error at load")

            ''Error Checking
            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()


                MsgBox("The connection to the P drive was interupted..@ store bttn click")


            End If




        Catch ex As Exception


            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try





    End Sub




    Private Sub radNewHire_CheckedChanged(sender As Object, e As EventArgs) Handles radNewHire.CheckedChanged


        Me.txtEmail.TabIndex = 4



        Button8.Enabled = False
        Button8.Visible = False


        btnFind.Visible = True
        btnFind.Enabled = True

        Button7.Enabled = False
        Button7.Visible = False





        lblEmail.Visible = True
        txtEmail.Visible = True


        ''
        Me.txtEmail.Location = New System.Drawing.Point(119, 209)
        Me.lblEmail.Location = New System.Drawing.Point(71, 216)



        lblEmpID.Text = "User ID:"
        lblEmpID.Location = New System.Drawing.Point(13, 224)

        ''
        Me.txtEmail.Location = New System.Drawing.Point(119, 209)

        If lblEmpID.Visible = True Then
            lblEmpID.Visible = False
        End If

        If txtUserID.Visible = True Then
            txtUserID.Visible = False
        End If


        If lblOtherEmailOption.Visible = True Then
            lblOtherEmailOption.Visible = False
        End If


        If txtOtherOptionTxt.Visible = True Then
            txtOtherOptionTxt.Visible = False

        End If

        If txtClientID.Visible = True Then
            txtClientID.Visible = False
        End If

        If lblClientID.Visible = True Then
            lblClientID.Visible = False
        End If


        If lblNewHireEmail.Visible = True Then

            lblNewHireEmail.Visible = False
        End If

        If txtNewHireEmail.Visible = True Then

            txtNewHireEmail.Visible = False
        End If


        ''

        If txtNewHireEmail.Visible = True Then

            txtNewHireEmail.Visible = False


        End If


        If txtApplicationID.Visible = True Then

            txtApplicationID.Visible = False

        End If

        If lblAppID.Visible = True Then

            lblAppID.Visible = False

        End If

        If txtOrderID.Visible = True Then

            txtOrderID.Visible = False

        End If

        If lblOrderID.Visible = True Then

            lblOrderID.Visible = False

        End If

        If lblAccountName.Visible = True Then

            lblAccountName.Visible = False

        End If

        If txtAccountName.Visible = True Then

            txtAccountName.Visible = False

        End If

        If txtAppName.Visible = True Then

            txtAppName.Visible = False
        End If

        If lblAppName.Visible = True Then

            lblAppName.Visible = False
        End If


    End Sub

    Private Sub radManager_CheckedChanged(sender As Object, e As EventArgs) Handles radManager.CheckedChanged




        Me.txtUserID.TabIndex = 4

        Me.txtNewHireEmail.TabIndex = 5

   

        Button8.Enabled = False
        Button8.Visible = False

        btnFind.Visible = True
        btnFind.Enabled = True

        Button7.Enabled = False
        Button7.Visible = False






        Me.txtNewHireEmail.Location = New System.Drawing.Point(119, 236)
        Me.lblNewHireEmail.Location = New System.Drawing.Point(68, 241)


        lblEmpID.Visible = True
        txtUserID.Visible = True
        txtNewHireEmail.Visible = True
        lblNewHireEmail.Visible = True
        txtNewHireEmail.Text = "NoEmail@NoEmail.com"
        lblEmpID.Text = "User ID:"
        lblEmpID.Location = New System.Drawing.Point(57, 216)

        Me.txtUserID.Location = New System.Drawing.Point(119, 209)

        lblAppName.Visible = True
        txtAppName.Visible = True

        Me.txtAppName.Location = New System.Drawing.Point(119, 262)
        Me.lblAppName.Location = New System.Drawing.Point(3, 269)






        If txtClientID.Visible = True Then
            txtClientID.Visible = False
        End If

        If lblClientID.Visible = True Then
            lblClientID.Visible = False
        End If




        If lblEmail.Visible = True Then
            lblEmail.Visible = False
        End If

        If txtEmail.Visible = True Then
            txtEmail.Visible = False

        End If

        If lblOtherEmailOption.Visible = True Then
            lblOtherEmailOption.Visible = False

        End If

        If txtOtherOptionTxt.Visible = True Then
            txtOtherOptionTxt.Visible = False
        End If


        If lblAppID.Visible = True Then

            lblAppID.Visible = False

        End If

        If txtApplicationID.Visible = True Then

            txtApplicationID.Visible = False

        End If

        If txtOrderID.Visible = True Then

            txtOrderID.Visible = False

        End If

        If lblOrderID.Visible = True Then

            lblOrderID.Visible = False

        End If



        If lblAccountName.Visible = True Then

            lblAccountName.Visible = False

        End If

        If txtAccountName.Visible = True Then

            txtAccountName.Visible = False

        End If

    End Sub

    Private Sub radFCT_CheckedChanged(sender As Object, e As EventArgs) Handles radContractor.CheckedChanged



        Me.txtEmail.TabIndex = 4
        Me.txtUserID.TabIndex = 5
        Me.txtOrderID.TabIndex = 6
        Me.txtAccountName.TabIndex = 7

        btnFind.Visible = True
        btnFind.Enabled = True


        Button8.Enabled = False
        Button8.Visible = False

        Button7.Enabled = False
        Button7.Visible = False

        Me.lblAccountName.Location = New System.Drawing.Point(9, 297)
        Me.txtAccountName.Location = New System.Drawing.Point(119, 290)

        Me.txtEmail.Location = New System.Drawing.Point(119, 209)
        Me.lblEmail.Location = New System.Drawing.Point(72, 216)

        Me.lblAppID.Location = New System.Drawing.Point(57, 243)
        lblAppID.Text = "User ID:"

        Me.txtUserID.Location = New System.Drawing.Point(119, 236)

        Me.lblAppName.Location = New System.Drawing.Point(3, 324)
        Me.txtAppName.Location = New System.Drawing.Point(119, 317)


        lblEmail.Visible = True
        txtEmail.Visible = True

        lblOrderID.Visible = True
        txtOrderID.Visible = True



        lblAppID.Visible = True
        '  txtApplicationID.Visible = True
        lblAccountName.Visible = True
        txtAccountName.Visible = True
        txtUserID.Visible = True

        txtAppName.Visible = True
        lblAppName.Visible = True




        If lblEmpID.Visible = True Then
            lblEmpID.Visible = False
        End If




        'If txtUserID.Visible = True Then
        'txtUserID.Visible = False
        '  End If


        If lblOtherEmailOption.Visible = True Then
            lblOtherEmailOption.Visible = False
        End If


        If txtOtherOptionTxt.Visible = True Then
            txtOtherOptionTxt.Visible = False

        End If

        If txtClientID.Visible = True Then
            txtClientID.Visible = False
        End If

        If lblClientID.Visible = True Then
            lblClientID.Visible = False
        End If


        If lblNewHireEmail.Visible = True Then

            lblNewHireEmail.Visible = False
        End If

        If txtNewHireEmail.Visible = True Then

            txtNewHireEmail.Visible = False
        End If


    End Sub

    Private Sub radOther_CheckedChanged(sender As Object, e As EventArgs) Handles radOther.CheckedChanged







        Me.txtUserID.TabIndex = 4

        Me.txtAccountName.TabIndex = 5

        Me.txtEmail.TabIndex = 6




        btnFind.Visible = True
        btnFind.Enabled = True

        Button8.Enabled = False
        Button8.Visible = False

        Button7.Enabled = False
        Button7.Visible = False

        ''Put Account Name in right place
        Me.txtAccountName.Location = New System.Drawing.Point(106, 244)
        Me.lblAccountName.Location = New System.Drawing.Point(-1, 251)


        ''Put Email in right Place
        ' Me.txtOtherOptionTxt.Location = New System.Drawing.Point(106, 271)
        Me.txtEmail.Location = New System.Drawing.Point(106, 271)
        Me.lblOtherEmailOption.Location = New System.Drawing.Point(61, 279)



        lblAccountName.Visible = True
        txtAccountName.Visible = True
        lblOtherEmailOption.Visible = True
        ' txtOtherOptionTxt.Visible = True
        txtEmail.Visible = True
        txtOtherOptionTxt.Visible = False
        lblEmpID.Visible = True
        txtUserID.Visible = True

        lblEmpID.Text = "User ID:"
        lblEmpID.Location = New System.Drawing.Point(47, 224)

        Me.txtUserID.Location = New System.Drawing.Point(106, 217)


        If lblEmail.Visible = True Then
            lblEmail.Visible = False
        End If

        'If txtEmail.Visible = True Then
        '    txtEmail.Visible = False
        'End If

        If txtOtherOptionTxt.Visible = True Then


            txtOtherOptionTxt.Text = False

        End If




        ''
        If lblNewHireEmail.Visible = True Then

            lblNewHireEmail.Visible = False
        End If


        If txtNewHireEmail.Visible = True Then

            txtNewHireEmail.Visible = False
        End If


        If lblAppID.Visible = True Then

            lblAppID.Visible = False

        End If

        If txtApplicationID.Visible = True Then

            txtApplicationID.Visible = False

        End If

        If txtOrderID.Visible = True Then

            txtOrderID.Visible = False

        End If

        If lblOrderID.Visible = True Then

            lblOrderID.Visible = False

        End If



    End Sub



    Private Sub radFingerPrinting_CheckedChanged(sender As Object, e As EventArgs) Handles radFingerPrinting.CheckedChanged





        Me.txtEmail.TabIndex = 4
        Me.txtAccountName.TabIndex = 5

        Button8.Enabled = False
        Button8.Visible = False

        btnFind.Visible = True
        btnFind.Enabled = True


        Button7.Enabled = False
        Button7.Visible = False


        ''Account Nmame 
        Me.lblAccountName.Location = New System.Drawing.Point(9, 244)
        Me.txtAccountName.Location = New System.Drawing.Point(119, 236)

        Me.txtEmail.Location = New System.Drawing.Point(119, 209)
        Me.lblEmail.Location = New System.Drawing.Point(72, 216)


        'lblEmpID.Text = "User ID:"
        'lblEmpID.Location = New System.Drawing.Point(47, 224)


        lblEmail.Visible = True
        txtEmail.Visible = True
        lblAccountName.Visible = True
        txtAccountName.Visible = True



        If lblEmpID.Visible = True Then
            lblEmpID.Visible = False
        End If

        If txtUserID.Visible = True Then
            txtUserID.Visible = False
        End If


        If lblOtherEmailOption.Visible = True Then
            lblOtherEmailOption.Visible = False
        End If


        If txtOtherOptionTxt.Visible = True Then
            txtOtherOptionTxt.Visible = False

        End If

        If txtClientID.Visible = True Then
            txtClientID.Visible = False
        End If

        If lblClientID.Visible = True Then
            lblClientID.Visible = False
        End If

        If lblNewHireEmail.Visible = True Then

            lblNewHireEmail.Visible = False
        End If


        If txtNewHireEmail.Visible = True Then

            txtNewHireEmail.Visible = False
        End If



        If lblAppID.Visible = True Then

            lblAppID.Visible = False

        End If

        If txtApplicationID.Visible = True Then

            txtApplicationID.Visible = False

        End If

        If txtOrderID.Visible = True Then

            txtOrderID.Visible = False

        End If

        If lblOrderID.Visible = True Then

            lblOrderID.Visible = False

        End If

        If txtAppName.Visible = True Then

            txtAppName.Visible = False
        End If

        If lblAppName.Visible = True Then

            lblAppName.Visible = False
        End If

    End Sub


    Private Sub radPROM_CheckedChanged(sender As Object, e As EventArgs) Handles radPROM.CheckedChanged




        Me.txtUserID.TabIndex = 4
        Me.txtEmail.TabIndex = 5


        Button7.Enabled = False
        Button7.Visible = False



        Button8.Enabled = False
        Button8.Visible = False

        btnFind.Visible = True
        btnFind.Enabled = True


        lblEmpID.Text = "Vendor ID:"
        lblEmpID.Location = New System.Drawing.Point(39, 216)
        Me.txtUserID.Location = New System.Drawing.Point(119, 209)



        Me.txtEmail.Location = New System.Drawing.Point(119, 236)

        Me.lblNewHireEmail.Location = New System.Drawing.Point(71, 243)


        lblEmpID.Visible = True
        txtUserID.Visible = True
        ' txtNewHireEmail.Visible = True
        txtEmail.Visible = True
        lblNewHireEmail.Visible = True
        '  txtNewHireEmail.Text = ""
        


      

        If txtClientID.Visible = True Then
            txtClientID.Visible = False
        End If

        If lblClientID.Visible = True Then
            lblClientID.Visible = False
        End If




        If lblEmail.Visible = True Then
            lblEmail.Visible = False
        End If

        'If txtEmail.Visible = True Then
        '    txtEmail.Visible = False

        'End If

        If txtNewHireEmail.Visible = True Then


            txtNewHireEmail.Visible = False

        End If



        If lblOtherEmailOption.Visible = True Then
            lblOtherEmailOption.Visible = False

        End If

        If txtOtherOptionTxt.Visible = True Then
            txtOtherOptionTxt.Visible = False
        End If


        If lblAppID.Visible = True Then

            lblAppID.Visible = False

        End If

        If txtApplicationID.Visible = True Then

            txtApplicationID.Visible = False

        End If

        If txtOrderID.Visible = True Then

            txtOrderID.Visible = False

        End If

        If lblOrderID.Visible = True Then

            lblOrderID.Visible = False

        End If



        If lblAccountName.Visible = True Then

            lblAccountName.Visible = False

        End If

        If txtAccountName.Visible = True Then

            txtAccountName.Visible = False

        End If

        If txtAppName.Visible = True Then

            txtAppName.Visible = False
        End If

        If lblAppName.Visible = True Then

            lblAppName.Visible = False
        End If





    End Sub


    Public Sub DoLogLaterstuff()

        ''======================================== IF the Call is already stored do this==============================================================================

        ' If lblSNoti.Visible = True Then

        '' LogTicket Counter / pending ticket info for log later box


        LogLaterCouner = LogLaterCouner + 1


        ' Log_Later.lblPendingTicketsNumber.Text = LogLaterCouner.ToString




        ''Call List Indicator

        LogLater_Timer.Enabled = True

        lblLogedLaterSTORED.Visible = True



        ''=========================== Call information getting sent to logLater Box================================================================================


        Log_Later.ListBox1.Items.Add(lblNewDate.Text + " / " + lblNewTime.Text)
        Log_Later.ListBox1.Items.Add("Call#: " + lblcCounter.Text)
        Log_Later.ListBox1.Items.Add("SPC Reference #: " + lblSPCRefNum.Text)

        If radNewHire.Checked = True Then

            Log_Later.ListBox1.Items.Add("CVS New Hire")

        ElseIf radManager.Checked = True Then


            Log_Later.ListBox1.Items.Add("CVS Store Manager")

        ElseIf radContractor.Checked = True Then

            Log_Later.ListBox1.Items.Add("Field Colleague Trainer")

        ElseIf radOther.Checked = True Then

            Log_Later.ListBox1.Items.Add("Warm Transfer")



        ElseIf radFingerPrinting.Checked = True Then

            Log_Later.ListBox1.Items.Add("FingerPrinting")

        ElseIf radClient.Checked = True Then

            Log_Later.ListBox1.Items.Add("Enterprise Client")


        ElseIf radApplicantEnt.Checked = True Then

            Log_Later.ListBox1.Items.Add("Enterprise Applicant")

        ElseIf radPROM.Checked = True Then

            Log_Later.ListBox1.Items.Add("PROM")


        End If



        Log_Later.ListBox1.Items.Add(txtFirstName.Text)
        Log_Later.ListBox1.Items.Add(txtLastName.Text)
        Log_Later.ListBox1.Items.Add(txtPhone.Text)

        If radNewHire.Checked = True Then

            Log_Later.ListBox1.Items.Add(txtEmail.Text)
            Log_Later.ListBox1.Items.Add("1-1134812539086")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")

        ElseIf radManager.Checked = True Then

            Log_Later.ListBox1.Items.Add(txtUserID.Text)
            Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)
            Log_Later.ListBox1.Items.Add("1-1134812539086")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")

            If txtAppName.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtAppName.Text)

            End If

        ElseIf radContractor.Checked = True Then

            Log_Later.ListBox1.Items.Add(txtUserID.Text)
            Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)
            Log_Later.ListBox1.Items.Add("1-1134812539086")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")


            If txtAppName.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtAppName.Text)

            End If

            '' Transfer ===================================================================================================================================
        ElseIf radOther.Checked = True Then


            If txtUserID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtUserID.Text)

            End If


            If txtClientID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtClientID.Text)

            End If


            If txtOtherOptionTxt.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtOtherOptionTxt.Text)

            End If

            If txtAccountName.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtAccountName.Text)

            End If

            Log_Later.ListBox1.Items.Add("N/a")


            '' FingerPrinting ================================================================================================================================================

        ElseIf radFingerPrinting.Checked = True Then


            If txtEmail.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")
            Else

                Log_Later.ListBox1.Items.Add(txtEmail.Text)

            End If


            If txtAccountName.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else

                Log_Later.ListBox1.Items.Add(txtAccountName.Text)

            End If

            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")

            '' Enterprise Client ===============================================================================================================================================
        ElseIf radClient.Checked = True Then


            If txtUserID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtUserID.Text)

            End If


            If txtClientID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtClientID.Text)

            End If




            If txtOrderID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtOrderID.Text)

            End If



            If txtEmail.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtEmail.Text)

            End If


            If txtAccountName.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtAccountName.Text)

            End If

            If txtAppName.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtAppName.Text)

            End If


            '' EnterPrise Applicant ============================================================================================================================================

        ElseIf radApplicantEnt.Checked = True Then



            If txtEmail.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtEmail.Text)

            End If

            If txtApplicationID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtApplicationID.Text)

            End If



            If txtOrderID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtOrderID.Text)

            End If


            Log_Later.ListBox1.Items.Add(txtAccountName.Text)
            Log_Later.ListBox1.Items.Add("N/a")


            '===========PROM==============================================================================================================================================

        ElseIf radPROM.Checked = True Then

            If txtUserID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else

                Log_Later.ListBox1.Items.Add(txtUserID.Text)

            End If

            If txtNewHireEmail.Text = "" Then


                Log_Later.ListBox1.Items.Add("N/a")
            Else

                Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)

            End If




            Log_Later.ListBox1.Items.Add("1-1138742798839")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")

        End If

        Log_Later.ListBox1.Items.Add(txtCallDetail.Text)
        Log_Later.ListBox1.Items.Add("!!!!!!!!!!!!!!--- Log This Call--!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        Log_Later.ListBox1.Items.Add("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")




        Log_Later.Show()






    End Sub

    Public Sub clearitems()

        ''Then Clear items


        ''Then Clear
        lblReflabel.Visible = False
        lblSPCRefNum.Visible = False


        lblSPCRefNum.Text = "Generating..."


        lblOfficalReference.ForeColor = Color.Red


        Me.lblEmpID.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

        Me.lblFirstName.ForeColor = System.Drawing.Color.Black
        Me.lblLastN.ForeColor = System.Drawing.Color.Black
        Me.lblPhone.ForeColor = System.Drawing.Color.Black
        Me.lblEmpID.ForeColor = System.Drawing.Color.Black
        Me.lblEmail.ForeColor = System.Drawing.Color.Black
        Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
        Me.lblClientID.ForeColor = System.Drawing.Color.Black
        Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black
        Me.lblAppName.ForeColor = System.Drawing.Color.Black


        lblOrderID.ForeColor = Color.Black
        lblAppID.ForeColor = Color.Black
        lblAccountName.ForeColor = Color.Black



        ''focus back on  first name when new button cleared

        Me.ActiveControl = txtFirstName


        txtFirstName.Clear()
        txtLastName.Clear()
        txtPhone.Clear()
        txtEmail.Clear()
        txtUserID.Clear()
        txtOtherOptionTxt.Clear()
        txtCallDetail.Clear()
        txtClientID.Clear()
        txtNewHireEmail.Clear()
        txtApplicationID.Clear()
        txtOrderID.Clear()
        txtAccountName.Clear()
        txtAppName.Clear()

        '  ClearBackender()





        radNewHire.Checked = False
        radManager.Checked = False
        radContractor.Checked = False
        radOther.Checked = False
        radFingerPrinting.Checked = False
        lblSNoti.Visible = False
        radClient.Checked = False
        radApplicantEnt.Checked = False
        radPROM.Checked = False

        lblAppName.Visible = False
        txtAppName.Visible = False



        If lblEmpID.Visible = True Then
            lblEmpID.Visible = False
        End If

        If txtUserID.Visible = True Then
            txtUserID.Visible = False
        End If

        If lblEmail.Visible = True Then
            lblEmail.Visible = False

        End If

        If txtEmail.Visible = True Then
            txtEmail.Visible = False

        End If

        If lblOtherEmailOption.Visible = True Then
            lblOtherEmailOption.Visible = False
        End If

        If txtOtherOptionTxt.Visible = True Then

            txtOtherOptionTxt.Visible = False
        End If

        If lblClientID.Visible = True Then

            lblClientID.Visible = False
        End If

        If txtClientID.Visible = True Then

            txtClientID.Visible = False
        End If

        If lblOrderID.Visible = True Then

            lblOrderID.Visible = False

        End If

        If txtOrderID.Visible = True Then

            txtOrderID.Visible = False

        End If

        If lblAppID.Visible = True Then

            lblAppID.Visible = False


        End If

        If txtApplicationID.Visible = True Then

            txtApplicationID.Visible = False

        End If

        If lblAccountName.Visible = True Then

            lblAccountName.Visible = False

        End If

        If txtAccountName.Visible = True Then

            txtAccountName.Visible = False

        End If

        lblNewHireEmail.Visible = False
        txtNewHireEmail.Visible = False


        'cboPhoneRef.Text = "Phone Number Reference List"
        ''lblPhoneRefNum.Visible = False
        ''lblPhoneRefNum.ForeColor = Color.Blue


        lblLogedLaterSTORED.Visible = False

        lblCallList_Indicator.Visible = False


    End Sub





    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click

        Try

            ''make sure give program enough time to store call before hitt new button

            If lblCallList_Indicator.Visible = False And lblSNoti.Visible = True Then


                MessageBox.Show("Oops, button pressed before process fully completed..", "Warning", MessageBoxButtons.RetryCancel)


            End If

            If lblSNoti.Visible = False Then


                If MessageBox.Show("Be advised you are about to clear all fields, do you wish to proceed?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then

                    '' dont clear items


                Else

                    ''clear items if yes

                    ''Reset Siebel Connect Buttons 

                    ResetButtons()



                    lblReflabel.Visible = False
                    lblSPCRefNum.Visible = False


                    lblSPCRefNum.Text = "Generating..."


                    lblOfficalReference.ForeColor = Color.Red


                    Me.lblEmpID.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

                    Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                    Me.lblLastN.ForeColor = System.Drawing.Color.Black
                    Me.lblPhone.ForeColor = System.Drawing.Color.Black
                    Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                    Me.lblEmail.ForeColor = System.Drawing.Color.Black
                    Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                    Me.lblClientID.ForeColor = System.Drawing.Color.Black
                    Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black

                    Me.lblAppName.ForeColor = System.Drawing.Color.Black


                    lblOrderID.ForeColor = Color.Black
                    lblAppID.ForeColor = Color.Black
                    lblAccountName.ForeColor = Color.Black



                    ''focus back on  first name when new button cleared

                    Me.ActiveControl = txtFirstName


                    txtFirstName.Clear()
                    txtLastName.Clear()
                    txtPhone.Clear()
                    txtEmail.Clear()
                    txtUserID.Clear()
                    txtOtherOptionTxt.Clear()
                    txtCallDetail.Clear()
                    txtClientID.Clear()
                    txtNewHireEmail.Clear()
                    txtApplicationID.Clear()
                    txtOrderID.Clear()
                    txtAccountName.Clear()
                    txtAppName.Clear()

                    ClearBackender()





                    radNewHire.Checked = False
                    radManager.Checked = False
                    radContractor.Checked = False
                    radOther.Checked = False
                    radFingerPrinting.Checked = False
                    lblSNoti.Visible = False
                    radClient.Checked = False
                    radApplicantEnt.Checked = False
                    radPROM.Checked = False

                    lblAppName.Visible = False
                    txtAppName.Visible = False


                    If lblEmpID.Visible = True Then
                        lblEmpID.Visible = False
                    End If

                    If txtUserID.Visible = True Then
                        txtUserID.Visible = False
                    End If

                    If lblEmail.Visible = True Then
                        lblEmail.Visible = False

                    End If

                    If txtEmail.Visible = True Then
                        txtEmail.Visible = False

                    End If

                    If lblOtherEmailOption.Visible = True Then
                        lblOtherEmailOption.Visible = False
                    End If

                    If txtOtherOptionTxt.Visible = True Then

                        txtOtherOptionTxt.Visible = False
                    End If

                    If lblClientID.Visible = True Then

                        lblClientID.Visible = False
                    End If

                    If txtClientID.Visible = True Then

                        txtClientID.Visible = False
                    End If

                    If lblOrderID.Visible = True Then

                        lblOrderID.Visible = False

                    End If

                    If txtOrderID.Visible = True Then

                        txtOrderID.Visible = False

                    End If

                    If lblAppID.Visible = True Then

                        lblAppID.Visible = False


                    End If

                    If txtApplicationID.Visible = True Then

                        txtApplicationID.Visible = False

                    End If

                    If lblAccountName.Visible = True Then

                        lblAccountName.Visible = False

                    End If

                    If txtAccountName.Visible = True Then


                        txtAccountName.Visible = False

                    End If

                    lblNewHireEmail.Visible = False
                    txtNewHireEmail.Visible = False


                    'cboPhoneRef.Text = "Phone Number Reference List"
                    ''lblPhoneRefNum.Visible = False
                    ''lblPhoneRefNum.ForeColor = Color.Blue


                    lblLogedLaterSTORED.Visible = False

                    lblCallList_Indicator.Visible = False


                    lblExsitorNot.Text = "Durrant"

                End If




            End If

            If lblLogedLaterSTORED.Visible = True And lblSNoti.Visible = True Then


                Dim result1 As Integer = MessageBox.Show("Be advised you are about to clear all fields, do you wish to proceed?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If result1 = DialogResult.Yes Then


                    clearitems()

                    lblExsitorNot.Text = "Durrant"



                ElseIf result1 = DialogResult.No Then



                End If



            End If


            If lblSNoti.Visible = True And lblLogedLaterSTORED.Visible = False Then

                Dim result As Integer = MessageBox.Show("Has this call been logged yet?", "Scratch Pad Compliance", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)



                If result = DialogResult.No Then

                    DoLogLaterstuff()

                    clearitems()

                    lblExsitorNot.Text = "Durrant"

                ElseIf result = DialogResult.Yes Then

                    '' Count call
                    Log_Ticket_Counter()

                    lblTicketsLoggedNumber.Visible = True
                    lblDailyTicketLogCount.Visible = True


                    clearitems()

                    lblExsitorNot.Text = "Durrant"

                ElseIf result = DialogResult.Cancel Then


                End If



            End If


          
            






            ''Error Checking
        Catch ex As OleDbException

            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()




                MsgBox("The connection to the P drive was interupted..@ new button")


            End If

        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            MsgBox("system error at new button")




        Catch ex As Exception


            MsgBox(ex.Message, 0 Or 48, "Alert")




        End Try








    End Sub



    Private Sub btnLogLater_Click(sender As Object, e As EventArgs)


        Try


            Me.Cursor = Cursors.WaitCursor

            ''Make sure job title is selected

            If radManager.Checked = False And radNewHire.Checked = False And radOther.Checked = False And radContractor.Checked = False And radApplicantEnt.Checked = False And radClient.Checked = False Then

                MessageBox.Show("Please be advised that a Job title must be selected in order to save this call, please select a Job Title and try again", "Warning", MessageBoxButtons.RetryCancel)

                lblLogedLaterSTORED.Visible = False

                Me.Cursor = Cursors.Hand

            Else

                ''No Duplicate Log Laters

                If lblLogedLaterSTORED.Visible = True Then

                    MessageBox.Show("Please be advised this call has already been placed in the ‘Log Later’ box", "Warning", MessageBoxButtons.RetryCancel)


                    Me.Cursor = Cursors.Hand
                Else


                    '' Make sure all required fields filled in 

                    If txtFirstName.Text = "" Then

                        MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                        Me.ActiveControl = txtFirstName

                        lblLogedLaterSTORED.Visible = False

                        Me.Cursor = Cursors.Hand

                    Else

                        '' Make sure all required fields filled in 

                        If txtLastName.Text = "" Then

                            MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                            Me.ActiveControl = txtLastName

                            Me.Cursor = Cursors.Hand

                        Else

                            '' Make sure all required fields filled in 

                            If txtPhone.Text = "" Then


                                MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                                Me.ActiveControl = txtPhone

                                Me.Cursor = Cursors.Hand

                            Else


                                '' Make sure all required fields filled in 

                                If txtCallDetail.Text = "" Then

                                    MessageBox.Show("Please fill out 'Call Detail' section before using this option", "Warning", MessageBoxButtons.RetryCancel)

                                    Me.ActiveControl = txtCallDetail


                                    lblLogedLaterSTORED.Visible = False

                                    Me.Cursor = Cursors.Hand

                                Else

                                    ''If call was store already===============================================================================================================================================================================


                                    If lblSNoti.Visible = True Then

                                        LogLater()


                                        lblLogedLaterSTORED.Visible = True



                                        ''If ticket is NOT stored do this============================================================================================================================================


                                    ElseIf lblSNoti.Visible = False Then



                                        ''Call Counter Timer ( counts call)
                                        CallCounterTimer.Enabled = True



                                        ''Store Call Thread ( stores call to database)


                                        LoglaterStoreCall = New System.Threading.Thread(AddressOf LogStoreCall)

                                        LoglaterStoreCall.Start()



                                        ''Call List Timer( sends call to call list)
                                        LogLaterCallListTimer.Enabled = True


                                        ''Call List Indicator

                                        LogLater_Timer.Enabled = True





                                        ''========== Turn Labels Blue ===========================================================================================================================================

                                        Me.lblFirstName.ForeColor = Color.Blue
                                        Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                                        Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                                        Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                                        Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue

                                        Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                                        Me.lblClientID.ForeColor = System.Drawing.Color.Blue

                                        Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue

                                        lblOrderID.ForeColor = Color.Blue
                                        lblAppID.ForeColor = Color.Blue

                                        lblSNoti.Visible = True

                                        Me.Cursor = Cursors.Hand

                                        lblDailyCallC.Visible = True
                                        lblcCounter.Visible = True


                                        lblReflabel.Visible = True
                                        lblSPCRefNum.Visible = True

                                        lblLogedLaterSTORED.Visible = True



                                    End If




                                    Me.Cursor = Cursors.Hand

                                End If




                            End If




                        End If
                    End If
                End If
            End If



        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")


        End Try




    End Sub

    Public Sub Template_Password()

        '' Password Reset Templates

        'CVS----------------------------------------


        If ComboBox1.SelectedItem = "CVS Manager Password Reset" Then

            txtScrapeBoxTitle.Text = "CVS Manager Password Reset"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                            "Verified Store #, and Store Address" & vbCrLf &
                                "User needed a password reset" & vbCrLf &
          "Advised User to find default password in LearnNet/ User Knew Default Password" & vbCrLf &
                                "Reset password to default" & vbCrLf &
                            "User was successfully able to log in" & vbCrLf &
                        "Offered additional assistance & Customer Declined" & vbCrLf &
                                        "Closed Ticket"

            If lblScrapedNoti.Visible = True Then

                lblScrapedNoti.Visible = False


            End If

        End If

        If ComboBox1.SelectedItem = "New Applicant Password Reset- Resent Credentials" Then


            txtScrapeBoxTitle.Text = "New Applicant Password Reset- Resent Credentials"

            txtScrapeBox.Text = "Obtained Name, and Phone #" & vbCrLf &
            "Verified SSN, and Email address, DOB" & vbCrLf &
            "Applicant could not log on/ locked out" & vbCrLf &
            "Resent login emails – received" & vbCrLf &
            "Walked applicant through password reset" & vbCrLf &
            "Offered additional assistance & Customer Declined" & vbCrLf &
            "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "New Applicant Password Reset - Created Custom Password" Then

            txtScrapeBoxTitle.Text = "New Applicant Password Reset - Created Custom Password"

            txtScrapeBox.Text = "Obtained Name, and Phone #" & vbCrLf &
            "Verified SSN, and Email address, DOB" & vbCrLf &
            "Applicant could not log in with credentials sent to them in the email" & vbCrLf &
            "Advised applicant i would create them a custom log in name and password" & vbCrLf &
            "Created applicant log in credentials" & vbCrLf &
            "Applicant was successfully able to log in" & vbCrLf &
            "Offered additional assistance & Customer Declined" & vbCrLf &
            "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If






    End Sub

    Public Sub Template_ApplicantStatusOfEmployment()

        ''CVS--------------------

        If ComboBox1.SelectedItem = "Archival Paper I-9 - New I-9 Needed" Then

            txtScrapeBoxTitle.Text = "Archival Paper I-9 - New I-9 Needed"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                "Verified Store # and Address" & vbCrLf &
                "User called in with a question about a Rehire with a Paper I-9 on file" & vbCrLf &
                "Advised User that I will send out a request to CVS HR to have a new I-9 created" & vbCrLf &
                "Advised user that it may take up to 48 hours for the request to be resolved" & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.Text = "Applicant with Incorrect Email – Resent Login Credentials" Then

            txtScrapeBoxTitle.Text = "Applicant with Incorrect Email – Resent Login Credentials"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                                "Verified DOB/Home Address, and SSN" & vbCrLf &
                                "User called in and did not receive emails" & vbCrLf &
                                "Determined incorrect email listed in Guardian" & vbCrLf &
                                "Advised applicant I will resend log in credentials to new email address" & vbCrLf &
                                "Resent Emails" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.SelectedItem = "Approval of I-9 Request" Then

            txtScrapeBoxTitle.Text = "Approval of I-9 Request"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                "Verified DOB, last 4 digits of SSN, and Email Address" & vbCrLf &
                "Applicant wanted to check on the status of their i-9" & vbCrLf &
                "Informed applicant that the i-9 was completed bu is waiting to be approved by HR " & vbCrLf &
                "Informed applicant that the i-9 was completed bu is waiting to be approved by HR " & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



        If ComboBox1.SelectedItem = "I-9 Status Check By Applicant" Then

            txtScrapeBoxTitle.Text = "I-9 Status Check By Applicant"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
             "Verified DOB, SSN And email" & vbCrLf &
             "User called to check status of I-9" & vbCrLf &
             "Advised applicant that Section 1 of I-9 Completed – Not Completed" & vbCrLf &
             "Offered additional assistance & Customer Declined" & vbCrLf &
             "Closed Ticket"


        End If


        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False

        End If


        If ComboBox1.SelectedItem = "I-9 Status Check By Manager" Then

            txtScrapeBoxTitle.Text = "I-9 Status Check By Manager"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
             "Verified Store #, and Store Address" & vbCrLf &
             "User called to check status of I-9 for an employee" & vbCrLf &
             "Employee:" & vbCrLf &
             "Advised User that Section 1 of I-9 Completed – Not Completed" & vbCrLf &
             "Offered additional assistance & Customer Declined" & vbCrLf &
             "Closed Ticket"



        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False

        End If


        If ComboBox1.SelectedItem = "Rehire Needs New I-9 Created – Applicant Called In" Then

            txtScrapeBoxTitle.Text = "Rehire Needs New I-9 Created – Applicant Called In"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                                "Verified DOB, last 4 digits of SSN, and Email Address" & vbCrLf &
                                "Applicant Called in With Question about I-9" & vbCrLf &
                                "Determined Applicant is a Rehire" & vbCrLf &
                                "Advised Applicant that i will send out a request to HR to Create a New I-9" & vbCrLf &
                                "Advised Rehire to monitor email for new log in credentials " & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"


        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "Rehire Needs New I-9 Created – Manager Called In" Then

            txtScrapeBoxTitle.Text = "Rehire Needs New I-9 Created – Manager Called In"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                                "Verified Store # and Address" & vbCrLf &
                                "Manager Stated a Rehire Needs a New I-9" & vbCrLf &
                                "Employee:" & vbCrLf &
                                "Advised Manager I will send out a request to HR to have new I-9 created" & vbCrLf &
                                "Advised Manager it may take up to 48 hours for HR to respond" & vbCrLf &
                                "Advised Manager to have Rehire monitor their email for new log in credentials " & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"


        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.SelectedItem = "Wrong SSN on I-9 – Applicant Called In" Then

            txtScrapeBoxTitle.Text = "Wrong SSN on I-9 – Applicant Called In"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                                "Verified DOB, Last Four of SSN, and Email" & vbCrLf &
                                "Applicant Stated their SSN was entered incorrect on I-9" & vbCrLf &
                                "Determined SSN was incorrect" & vbCrLf &
                                "Advised applicant that they need to take the incorrect SSN to their manager or FCT" & vbCrLf &
                                "Advised applicant that the manager or FCT will use the incorrect SSN to locate them and have them redo section 1" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"


        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "Wrong SSN on I-9 – Manager / FCT Called In" Then

            txtScrapeBoxTitle.Text = "Wrong SSN on I-9 – Manager / FCT Called In"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                                "Verified Store #, and Store Address" & vbCrLf &
                                "User called in and could not locate emp with SSN" & vbCrLf &
                                "Employee:" & vbCrLf &
                                "Determined SSN was incorrect" & vbCrLf &
                                "Advised manager that she could have employee redo section 1 of i-9 to update SSN" & vbCrLf &
                                "Successfully walked Manager through the process and changed SSN on i-9  " & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"


        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If








    End Sub

    Public Sub Template_ProductNavigation()


        If ComboBox1.SelectedItem = "Applicant Stuck on Step 3 of I-9" Then

            txtScrapeBoxTitle.Text = "Applicant Stuck on Step 3 of I-9"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                                "Verified DOB, last 4 digits of SSN, and Email Address" & vbCrLf &
                                "Applicant stated they are stuck on section 3 and cannot continue to section 4" & vbCrLf &
                                "Advised applicant that section 1 was completed and that is all they are responsible for" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"


        End If


        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.SelectedItem = "Advising New Hire how to Redo Sec 1" Then

            txtScrapeBoxTitle.Text = "Advising New Hire how to Redo Sec 1"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                                    "Verified DOB, SSN And email" & vbCrLf &
                                     "Applicant Stated they did something wrong on sec 1 of I-9" & vbCrLf &
                                    "Applicant informed me that they was not in the store or at training class" & vbCrLf &
                                    "Advised applicant that manager must initiate the process that would allow them to redo section 1" & vbCrLf &
                                    "Advised applicant to have manager or HR call us and we can walk them through the process" & vbCrLf &
                                    "Offered additional assistance & Customer Declined" & vbCrLf &
                                    "Closed ticket"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.SelectedItem = "Creating New I-9 Assistance" Then

            txtScrapeBoxTitle.Text = "Creating New I-9 Assistance"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                                    "Verified Store # and Address" & vbCrLf &
                                     "User called in with a question about a New Hire / Rehire" & vbCrLf &
                                    "Determined a new I-9 had to be created" & vbCrLf &
                                    "Advised User how to create a new I-9" & vbCrLf &
                                    "I-9 Successfully created" & vbCrLf &
                                    "Offered additional assistance & Customer Declined" & vbCrLf &
                                    "Closed ticket"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If





        If ComboBox1.SelectedItem = "CVS W-4 Edits - MyHR" Then



            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                            "Verified SSN, DOB, and Email address" & vbCrLf &
                            "Applicant called in with a question regarding to StarSouce" & vbCrLf &
                            "Explained to applicant the difference between i-9 and W-4" & vbCrLf &
                            "Provided MYHR Number to applicant" & vbCrLf &
                            "Provided phone number: 1-888-694-7287" & vbCrLf &
                            "Offered additional assistance & Customer Declined" & vbCrLf &
                                              "Closed Ticket"


            txtScrapeBoxTitle.Text = "CVS W-4 Edits - MyHR"


        End If




        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If






        If ComboBox1.SelectedItem = "CVS OnDocs Assistance" Then



            txtScrapeBox.Text = "Obtained Name, Phone #, Client ID" & vbCrLf &
                            "Verified Email & User ID" & vbCrLf &
                            "User was having trouble trying to upload List A document to Guardian" & vbCrLf &
                            "Asked FCT if he or she had another Web-Browser - FCT declined" & vbCrLf &
                            "Advised FCT to fax the document to HR, so they can complete it on their end" & vbCrLf &
                            "Provided fax number: 401-653 -1119" & vbCrLf &
                            "Offered additional assistance & Customer Declined" & vbCrLf &
                                              "Closed Ticket"


            txtScrapeBoxTitle.Text = "CVS OnDocs Assistance"


        End If




        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



        If ComboBox1.SelectedItem = "Editing Section 1 – Applicant" Then

            txtScrapeBoxTitle.Text = "Editing Section 1 – Applicant"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                "Verified DOB, last 4 digits of SSN, and Email Address" & vbCrLf &
                "Applicant stated they put some incorrect information on section 1 of I-9" & vbCrLf &
                "Pulled up applicant and saw that only section 1 was completed" & vbCrLf &
                "Advised applicant that their manager needs to initiate that process to have them edit section 1" & vbCrLf &
                "Advised applicant to have the manager to give us a call" & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



        If ComboBox1.SelectedItem = "Editing Section 1 – Manager / FCT" Then

            txtScrapeBoxTitle.Text = "Editing Section 1 – Manager / FCT"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                "Verified Store # and Address" & vbCrLf &
                "User needed assistance with initiating the process to edit section 1 of I-9" & vbCrLf &
                "Advised manager how to initiate the process" & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If




        If ComboBox1.SelectedItem = "General I-9 Completion Assistance - Applicant" Then

            txtScrapeBoxTitle.Text = "General I-9 Completion Assistance - Applicant"

            txtScrapeBox.Text = "Obtained Name, Phone # " & vbCrLf &
                                "Verified DOB, SSN And email" & vbCrLf &
                                "Applicant needed help with completing section 1 of an I-9" & vbCrLf &
                                "Walked applicant step by step with completing section 1 of an I-9" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If




        If ComboBox1.SelectedItem = "General I-9 Completion Assistance - Manager" Then

            txtScrapeBoxTitle.Text = "General I-9 Completion Assistance - Manager"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                                "Verified Store #, and Store Address" & vbCrLf &
                                "User needed help with completing section 2 of an I-9" & vbCrLf &
                                "Walked User step by step with completing section 2 of the I-9" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If




        If ComboBox1.SelectedItem = "Golden Corral - Manager" Then

            txtScrapeBoxTitle.Text = "Golden Corral - Manager"

            txtScrapeBox.Text = "Obtained Name, Phone # " & vbCrLf &
                                "Verified Store # and Store address" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If




        If ComboBox1.SelectedItem = "Product Navigation – referred to StarSource" Then

            txtScrapeBoxTitle.Text = "Product Navigation – referred to StarSource"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                "Verified SSN, DOB, and Email address" & vbCrLf &
                "Applicant called in with a question regarding to StarSouce" & vbCrLf &
                "Explained to applicant the difference between Guardian and StarSouce" & vbCrLf &
                "Looked up applicant in Guardian – Listed / Not Listed" & vbCrLf &
                "Provided StarSouce Support Number to applicant" & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "Pay Roll Request" Then

            txtScrapeBoxTitle.Text = "Pay Roll Request"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                "Verified Store # and Address" & vbCrLf &
                "User called in regarding an applicant who has not yet dropped into their payroll system" & vbCrLf &
                "Employee:" & vbCrLf &
                "Pulled up applicant in Guardian and saw that the I-9 just had to be approved by HR" & vbCrLf &
                "Advised User that HR provided us with an email to give to any manager needing to get an applicant sent to payroll" & vbCrLf &
                "Provided the email HRSS_StarSourceReqs@CvsCareMark.com" & vbCrLf &
                "Advised user that it may take up to 48 hours for the request to be resolved" & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If




    End Sub


    Public Sub Template_Account_Config()

        If ComboBox1.SelectedItem = "Manager Location Update" Then

            txtScrapeBoxTitle.Text = "Manager Location Update"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID " & vbCrLf &
            "Verified Store #, and Store Address" & vbCrLf &
            "User called stating he could not locate an employee in his To-Do list" & vbCrLf &
            "Determined user needed to get his account location updated" & vbCrLf &
            "Advised user that I will contact HR to have account location updated" & vbCrLf &
            "Advised user that it may take up to 48 hours for the request to be resolved" & vbCrLf &
            "Offered additional assistance & Customer Declined" & vbCrLf &
            "Closed Ticket"


        End If


        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False
        End If

        If ComboBox1.SelectedItem = "Employee Location Update" Then

            txtScrapeBoxTitle.Text = "Employee Location Update"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
            "Verified DOB,Last four of SSN, and Email" & vbCrLf &
            "Employee called in with an issue" & vbCrLf &
            "Determined emloyee needed to get his account location updated" & vbCrLf &
            "Advised employee that I will contact HR to have account location updated" & vbCrLf &
            "Advised employee that it may take up to 48 hours for the request to be resolved" & vbCrLf &
            "Offered additional assistance & Customer Declined" & vbCrLf &
            "Closed Ticket"


        End If


        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False
        End If



        If ComboBox1.SelectedItem = "Hire Date Update" Then

            txtScrapeBoxTitle.Text = "Hire Date Update"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                "Verified Store # and Address" & vbCrLf &
                "User called in with a question regarding updating the Hire Date for an applicant" & vbCrLf &
                "Advised User I will send a request to CVS HR to have Hire date Updated" & vbCrLf &
                "Advised user that it may take up to 48 hours for the request to be resolved" & vbCrLf &
                "Sent out Request" & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If




    End Sub

    Public Sub ResetButtons()



        btnExsit.Enabled = False

        btnNewContact.Enabled = False

        '  btnSR.Enabled = False

        '  btnLogTicketButton.Enabled = False

        ' BtnSSave.Enabled = False






    End Sub






    Public Sub Template_NewUser_Request()

        If ComboBox1.SelectedItem = "New Applicant Setup" Then

            txtScrapeBoxTitle.Text = "New Applicant Setup"

            txtScrapeBox.Text = "Obtained Name and Phone #" & vbCrLf &
                                "CVS new hire stated they did not get login info" & vbCrLf &
                                "Checked Guardian – not listed" & vbCrLf &
                                "Asked new hire when Star Source Onboarding was completed" & vbCrLf &
                                "New Hire informed me that it passed 48 hours since completion of Star Source" & vbCrLf &
                                "Advised new hire that I will send an email to CVS HR to get them manually added" & vbCrLf &
                                "Advised new hire to continue to monitor email for login info " & vbCrLf &
                                "Advised new hire that it takes up to 48 for HR to respond" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "New User Setup" Then

            txtScrapeBoxTitle.Text = "New User Setup"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                                "Verified Store #, and Store Addres" & vbCrLf &
                                "New User could not login" & vbCrLf &
                                "Checked Guardian – not listed" & vbCrLf &
                                "Advised User that I will email HRSS to get them manually added" & vbCrLf &
                                "Advised User that it is a 48 hour turnaround time" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"


        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.SelectedItem = "Manager Calling on Behalf of Employee – Not Listed" Then


            txtScrapeBoxTitle.Text = "Manager Calling on Behalf of Employee – Not Listed"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
            "Verified Store #, and Store Address" & vbCrLf &
             "Manager called to check status of employee" & vbCrLf &
            "Employee:" & vbCrLf &
            "Checked Guardian employee not listed" & vbCrLf &
            "Informed manager I will send a request to HR to have them added to Guardian" & vbCrLf &
            "Informed manager that it was take up to 48 hours for request to be resolved" & vbCrLf &
            "Informed manager to have new employee to monitor email for log in credentials" & vbCrLf &
            "Offered additional assistance & Customer Declined" & vbCrLf &
            "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



    End Sub


    Public Sub Template_HangUp_Transfer()



        If ComboBox1.SelectedItem = "Hang Up – No Assistance Needed" Then

            txtScrapeBoxTitle.Text = "Hang Up – No Assistance Needed"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                "Verified SSN, DOB, and Email address" & vbCrLf &
                "Applicant realized they did not need any more assistance/call dropped" & vbCrLf &
                "Tried to Call Back, No Answer" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.SelectedItem = "Warm Transfer - Drug Testing" Then

            txtScrapeBoxTitle.Text = "Warm Transfer"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
             "Verified DOB, SSN And email" & vbCrLf &
             "Customer Called in with an issue that was not related I-9" & vbCrLf &
             "Warm Transferred customer to the Drug Testing department " & vbCrLf &
             "Offered additional assistance & Customer Declined" & vbCrLf &
             "Closed Ticket"


        End If


        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False

        End If


        If ComboBox1.SelectedItem = "Warm Transfer - Drug Testing" Then

            txtScrapeBoxTitle.Text = "Warm Transfer"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
             "Verified DOB, SSN And email" & vbCrLf &
             "Customer Called in with an issue that was not related I-9" & vbCrLf &
             "Warm Transferred customer to the Drug Testing department " & vbCrLf &
             "Offered additional assistance & Customer Declined" & vbCrLf &
             "Closed Ticket"


        End If


        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False

        End If


        If ComboBox1.SelectedItem = "Customer concerning background check or I-9 assistance" Then


            txtScrapeBoxTitle.Text = "Customer concerning background check or I-9 assistance"

            txtScrapeBox.Text = "Obtained Name, Phone #, Client ID" & vbCrLf &
                                 "Verified User ID, Fax#, and Email" & vbCrLf &
                                "CU called regarding a background" & vbCrLf &
                                "CID#" & vbCrLf &
                                "Transferred" & vbCrLf &
                                "Closed ticket"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "Candidate concerning background check" Then

            txtScrapeBoxTitle.Text = "Candidate concerning background check"

            txtScrapeBox.Text = "Obtained Name, Phone #, and CID or Order ID" & vbCrLf &
                                    "User called regarding background" & vbCrLf &
                                    "Transferred" & vbCrLf &
                                    "Closed ticket"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



    End Sub

    Public Sub Template_Everify_Status()


        If ComboBox1.SelectedItem = "SSA TNC – Manager" Then

            txtScrapeBoxTitle.Text = "SSA TNC – Manager"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                "Verified Store # and Address" & vbCrLf &
                "Manager called in regarding an applicant who received an SSA TNC on I-9" & vbCrLf &
                "Explained to the Manager what an SSA TNC is and the process they have to take in order to fix it." & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



        If ComboBox1.SelectedItem = "SSA TNC – Applicant" Then

            txtScrapeBoxTitle.Text = "SSA TNC – Applicant"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                "Verified DOB, last 4 digits of SSN, and Email Address" & vbCrLf &
                "Applicant called in stating they received a SSA TNC" & vbCrLf &
                "Explained to the applicant what an SSA TNC is and the process they have to take in order to fix it." & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



        If ComboBox1.SelectedItem = "DHS TNC – Manager" Then

            txtScrapeBoxTitle.Text = "DHS TNC – Manager"

            txtScrapeBox.Text = "Obtained Name, Phone #, Employee ID" & vbCrLf &
                "Verified Store # and Address" & vbCrLf &
                "Manager called in regarding an applicant who received an DHS TNC on I-9" & vbCrLf &
                "Explained to the Manager what an DHS TNC is and the process they have to take in order to fix it." & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



        If ComboBox1.SelectedItem = "DHS TNC – Applicant" Then

            txtScrapeBoxTitle.Text = "DHS TNC – Applicant"

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
                "Verified DOB, last 4 digits of SSN, and Email Address" & vbCrLf &
                "Applicant called in stating they received a DHS TNC" & vbCrLf &
                "Explained to the applicant what an DHS TNC is and the process they have to take in order to fix it." & vbCrLf &
                "Offered additional assistance & Customer Declined" & vbCrLf &
                "Closed Ticket"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



    End Sub


    Public Sub Template_Emails()

        If ComboBox1.SelectedItem = "Email: Add New USER to LawLogix Guardian System" Then



            txtScrapeBox.Text = "Good Morning / Good Afternoon" & vbCrLf &
"Can you please add the following new user into the LawLogix Guardian system?  They are not currently listed" & vbCrLf &
"Name:" & vbCrLf &
"Phone:" & vbCrLf &
"Employee ID:" & vbCrLf &
"Store#/ Location:" & vbCrLf &
"Job title:" & vbCrLf &
"Thank You!"

            txtScrapeBoxTitle.Text = "Please Add New USER to LawLogix Guardian System"

            If lblScrapedNoti.Visible = True Then

                lblScrapedNoti.Visible = False


            End If


        End If

        If ComboBox1.SelectedItem = "Email: Add New Hire to LawLogix Guardian System" Then

            txtScrapeBox.Text = "Good Morning / Good Afternoon" & vbCrLf &
"Can you please add the following new hire into the LawLogix Guardian system?  They are not currently listed" & vbCrLf &
"Name:" & vbCrLf &
"Email:" & vbCrLf &
"Phone Number:" & vbCrLf &
"Thank You!"

            txtScrapeBoxTitle.Text = "Please Add New Hire to LawLogix Guardian System"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.Text = "Email: Update New Hire Email Address" Then

            txtScrapeBox.Text = "Good Morning / Good Afternoon" & vbCrLf &
"Can you please change the following new hire’s email address into the LawLogix Guardian system?" & vbCrLf &
"Name:" & vbCrLf &
"Employee ID:" & vbCrLf &
"Old Email:" & vbCrLf &
"New Email:" & vbCrLf &
"Thank You!"


            txtScrapeBoxTitle.Text = "Please Update New Hire Email Address in LawLogix Guardian System"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



        If ComboBox1.SelectedItem = "Email: Please Update User Store Location" Then


            txtScrapeBox.Text = "Good Morning / Good Afternoon" & vbCrLf &
"Can you please update the following User store location in the LawLogix Guardian system?" & vbCrLf &
"Name:" & vbCrLf &
"Employee ID:" & vbCrLf &
"Old Store#:" & vbCrLf &
"New Store#:" & vbCrLf &
"Job title:" & vbCrLf &
"Thank You!"

            txtScrapeBoxTitle.Text = "Please Update User Store Location in LawLogix Guardian System"



        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "Email: Update New Hire SSN" Then

            txtScrapeBox.Text = "Good Morning / Good Afternoon" & vbCrLf &
"Can you please update the following new hire SSN in the LawLogix Guardian system?" & vbCrLf &
"Name:" & vbCrLf &
"Employee ID:" & vbCrLf &
"Wrong SSN#:" & vbCrLf &
"Correct SSN#:" & vbCrLf &
"Thank You!"


            txtScrapeBoxTitle.Text = "Please Update New Hire SSN in LawLogix Guardian System"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "Email: Edit New Hire Name in Guardian" Then

            txtScrapeBox.Text = "Good Morning / Good Afternoon" & vbCrLf &
"Can you please edit the following new hire’s name in the LawLogix Guardian system?  They currently have it misspelled in the Guardian System." & vbCrLf &
"Incorrect Name:" & vbCrLf &
"Employee ID:" & vbCrLf &
"Updated Name:" & vbCrLf &
"Thank You!"


            txtScrapeBoxTitle.Text = "Please Edit New Hire Name in LawLogix Guardian System"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.SelectedItem = "Email: Remove Duplicate I-9 for Employee" Then

            txtScrapeBox.Text = "Good Afternoon / Good Morning" & vbCrLf &
    "Can you please update this applicant’s account?  This employee has two I-9s on file and one needs to be removed." & vbCrLf &
"Name:" & vbCrLf &
"Employee ID" & vbCrLf &
"I-9 Created on ------ needs to be removed." & vbCrLf &
"Name:" & vbCrLf &
"Thank You!"

            txtScrapeBoxTitle.Text = "Please Remove Duplicate I-9 for Employee"

        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "Email: Rehire - Create New I-9 in Guardian" Then

            txtScrapeBox.Text = "Good Morning / Good Afternoon" & vbCrLf &
          "Can you please create a new I-9 for the following Rehire hire in the LawLogix Guardian system?" & vbCrLf &
          "Name:" & vbCrLf &
          "Phone:" & vbCrLf &
          "Email:" & vbCrLf &
          "Thank You!"


            txtScrapeBoxTitle.Text = "Please Create a New I-9 for This Re-Hire in the LawLogix Guardian System"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


    End Sub


    Public Sub Template_Fingerprint()

        If ComboBox1.SelectedItem = "Collections Site Assistance" Then

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
          "Operator could not pull up candidate" & vbCrLf &
          "Pulled up candidate in Support reports" & vbCrLf &
          "Ensured that information was entered correctly" & vbCrLf &
          "Directed operator to account orders" & vbCrLf &
          "Operator was able to bring up candidate" & vbCrLf &
          "Offered additional assistance & Customer Declined" & vbCrLf &
          "Closed Ticket"


            txtScrapeBoxTitle.Text = "Collections Site Assistance - Fingerprinting"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "General Assistance" Then

            txtScrapeBox.Text = "Obtained Name, Phone #" & vbCrLf &
          "User needed Assistance with" & vbCrLf &
          "Advised operator how to" & vbCrLf &
          "User successfully" & vbCrLf &
          "Offered additional assistance & Customer Declined" & vbCrLf &
          "Closed Ticket"


            txtScrapeBoxTitle.Text = "General Assistance - Fingerprinting"
        End If

        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


    End Sub

    Public Sub Template_Enterprise()

        If ComboBox1.SelectedItem = "I-9 E-Verify" Then



            txtScrapeBox.Text = "Obtained Name, Phone #, Client ID" & vbCrLf &
                            "Verified Email & User ID" & vbCrLf &
                                "User called questioning E-verify Status" & vbCrLf &
          "Informed user to close case and start new i-9" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                            "Advised Customer to log out of system when finished" & vbCrLf &
                                              "Closed Ticket"


            txtScrapeBoxTitle.Text = "I-9 E-Verify"


        End If




        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If

        If ComboBox1.SelectedItem = "Closing a Case" Then



            txtScrapeBox.Text = "Obtained Name, Phone #, Client ID" & vbCrLf &
                            "Verified Email & User ID" & vbCrLf &
                                "User called requesting to have CID:  marked as decisional" & vbCrLf &
          "Contacted Level 2 and they assisted me with closing the element " & vbCrLf &
                                "Informed user case was closed as decisional " & vbCrLf &
                                 "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Advised user to log out when done" & vbCrLf &
                                              "Closed Ticket"

            txtScrapeBoxTitle.Text = "Closing a Case"

            If lblScrapedNoti.Visible = True Then

                lblScrapedNoti.Visible = False


            End If



        End If

        If ComboBox1.SelectedItem = "MVR Question" Then



            txtScrapeBox.Text = "Obtained Name, Phone #, Client ID" & vbCrLf &
                            "Verified Email & User ID" & vbCrLf &
                                "User misspelled candidate's last name " & vbCrLf &
          "Advised to reorder due to incorrect spelling" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                            "Advised Customer to log out of system when finished" & vbCrLf &
                                              "Closed Ticket"


            txtScrapeBoxTitle.Text = "MVR Question"


        End If




        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If



    End Sub

    Public Sub Template_PROM()



        If ComboBox1.SelectedItem = "Password Reset - PROM" Then



            txtScrapeBox.Text = "Obtained Name, Phone #, Client ID" & vbCrLf &
                                "Verified Email & User ID" & vbCrLf &
                                "User needed assistance with logging into PROM system" & vbCrLf &
                               "Reset User password back to same password / Created a New Password" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"


            txtScrapeBoxTitle.Text = "PROM - Password Reset"


        End If




        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If


        If ComboBox1.SelectedItem = "Java Issues - PROM" Then



            txtScrapeBox.Text = "Obtained Name, Phone #, Client ID" & vbCrLf &
                                "Verified Email & User ID" & vbCrLf &
                                "User needed assistance with configuring Java" & vbCrLf &
                               "Assisted User in Java configurations" & vbCrLf &
                                 "User was successfully able to log into PROM" & vbCrLf &
                                "Offered additional assistance & Customer Declined" & vbCrLf &
                                "Closed Ticket"


            txtScrapeBoxTitle.Text = "PROM - Java Configurations"


        End If




        If lblScrapedNoti.Visible = True Then

            lblScrapedNoti.Visible = False


        End If









    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged



        Template_Password()

        Template_ApplicantStatusOfEmployment()

        Template_ProductNavigation()

        Template_Account_Config()

        Template_NewUser_Request()

        Template_HangUp_Transfer()

        Template_Everify_Status()

        Template_Emails()

        Template_Fingerprint()

        Template_Enterprise()

        Template_PROM()



    End Sub

    Private Sub btnScrape_Click(sender As Object, e As EventArgs) Handles btnScrape.Click


        If txtScrapeBox.Text = "" Then

            MessageBox.Show("The ScrapeBox is empty there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

            Me.ActiveControl = txtScrapeBox

        Else




            ''
            lblScrapedNoti.Visible = True

            ''Scrape / Copy

            Clipboard.SetText(Me.txtScrapeBox.Text, TextDataFormat.Text)


            Me.lblTicketTitlebar.ForeColor = System.Drawing.Color.Black



        End If





    End Sub


    Private Sub btnShiftOver_Click(sender As Object, e As EventArgs) Handles btnShiftOver.Click



        txtCallDetail.Text &= Environment.NewLine & ""
        txtCallDetail.Text &= Environment.NewLine & txtScrapeBox.Text




    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        Try


            If MessageBox.Show("The title box and scrapebox will be cleared entirely are you sure you want to proceed?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then


            Else




                NEWBOX.Text = "Personal Templates"
                ComboBox2.Text = "CVS I-9 Templates"

                Me.lblTicketTitlebar.ForeColor = System.Drawing.Color.Black

                txtScrapeBox.Clear()
                txtScrapeBoxTitle.Clear()

                lblScrapedNoti.Visible = False


                RefeshCombo.Enabled = True



                lblScrapedNoti.Text = "Scraped"

                lblScrapedNoti.ForeColor = Color.Blue


                txtScrapeBoxTitle.BackColor = System.Drawing.SystemColors.ButtonHighlight

                txtScrapeBox.BackColor = System.Drawing.SystemColors.ButtonHighlight


            End If



            ''Error Checking
        Catch ex As OleDbException

            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()




                MsgBox("The connection to the P drive was interupted..@ clear button")


            End If






        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            MsgBox("system error at clear button ")




        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try





    End Sub

    Private Sub lblFirstName_Click(sender As Object, e As EventArgs) Handles lblFirstName.Click


        If txtFirstName.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

        Else



            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If




            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black

                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblClientID.ForeColor = System.Drawing.Color.Black

                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black
                lblAppName.ForeColor = Color.Black

                lblAccountName.ForeColor = Color.Black

                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black

            ElseIf lblSNoti.Visible = True Then



                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue

                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue
                lblAppName.ForeColor = Color.Blue

                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue

            End If

            Me.lblFirstName.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            Clipboard.SetText(Me.txtFirstName.Text, TextDataFormat.Text)

            lblScrapedNoti.Visible = False

            ''lblPhoneRefNum.ForeColor = Color.Blue


        End If



    End Sub

    Private Sub lblLastN_Click(sender As Object, e As EventArgs) Handles lblLastN.Click

        If txtLastName.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)


            Me.ActiveControl = txtLastName

        Else



            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If





            Clipboard.SetText(Me.txtLastName.Text, TextDataFormat.Text)




            If lblSNoti.Visible = False Then

                lblAppName.ForeColor = Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblClientID.ForeColor = System.Drawing.Color.Black
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black

                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black
                lblAccountName.ForeColor = Color.Black

            ElseIf lblSNoti.Visible = True Then



                lblAppName.ForeColor = Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue


                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue

            End If


            Me.lblLastN.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))








            lblScrapedNoti.Visible = False

            ''lblPhoneRefNum.ForeColor = Color.Blue


        End If




    End Sub

    Private Sub lblPhone_Click(sender As Object, e As EventArgs) Handles lblPhone.Click


        If txtPhone.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

            Me.ActiveControl = txtPhone

        Else

            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If





            Clipboard.SetText(Me.txtPhone.Text, TextDataFormat.Text)

            Me.lblPhone.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black

                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblClientID.ForeColor = System.Drawing.Color.Black

                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black

                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black
                lblAccountName.ForeColor = Color.Black
                lblAppName.ForeColor = Color.Black


            ElseIf lblSNoti.Visible = True Then


                lblAppName.ForeColor = Color.Blue
                Me.lblLastN.ForeColor = System.Drawing.Color.Blue

                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue

                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue


            End If

            lblScrapedNoti.Visible = False


            'lblPhoneRefNum.ForeColor = Color.Blue


        End If


    End Sub

    Private Sub lblEmpID_Click(sender As Object, e As EventArgs) Handles lblEmpID.Click


        If txtUserID.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)


            Me.ActiveControl = txtUserID

        Else

            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If




            Me.lblEmpID.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                Me.lblFirstName.ForeColor = System.Drawing.Color.Black

                Me.lblClientID.ForeColor = System.Drawing.Color.Black
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black


                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black
                lblAccountName.ForeColor = Color.Black
                lblAppName.ForeColor = Color.Black




            ElseIf lblSNoti.Visible = True Then


                lblAppName.ForeColor = Color.Blue
                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue

                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue

                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue

            End If

            Clipboard.SetText(Me.txtUserID.Text, TextDataFormat.Text)


            lblScrapedNoti.Visible = False


            'lblPhoneRefNum.ForeColor = Color.Blue

        End If



    End Sub

    Private Sub lblEmail_Click(sender As Object, e As EventArgs) Handles lblEmail.Click

        If txtEmail.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

            Me.ActiveControl = txtEmail


        Else

            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If





            Me.lblEmail.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            Clipboard.SetText(Me.txtEmail.Text, TextDataFormat.Text)

            If lblSNoti.Visible = False Then

                lblAppName.ForeColor = Color.Black
                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                '' Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblClientID.ForeColor = System.Drawing.Color.Black
                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black
                lblNewHireEmail.ForeColor = Color.Black
                lblAccountName.ForeColor = Color.Black



            ElseIf lblSNoti.Visible = True Then



                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                '' Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblClientID.ForeColor = System.Drawing.Color.Blue

                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
                lblNewHireEmail.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue
                lblAppName.ForeColor = Color.Blue


            End If


            lblScrapedNoti.Visible = False

            'lblPhoneRefNum.ForeColor = Color.Blue


        End If


    End Sub

    Private Sub lblNewHireEmail_Click(sender As Object, e As EventArgs) Handles lblNewHireEmail.Click

        If txtNewHireEmail.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

            Me.ActiveControl = txtNewHireEmail


        Else

            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If




            Me.lblNewHireEmail.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            Clipboard.SetText(Me.txtNewHireEmail.Text, TextDataFormat.Text)

            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblClientID.ForeColor = System.Drawing.Color.Black

                lblAccountName.ForeColor = Color.Black
                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black


            ElseIf lblSNoti.Visible = True Then



                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblClientID.ForeColor = System.Drawing.Color.Blue

                lblAccountName.ForeColor = Color.Blue
                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
            End If


            lblScrapedNoti.Visible = False

            'lblPhoneRefNum.ForeColor = Color.Blue


        End If



    End Sub

    Private Sub lblClientID_Click(sender As Object, e As EventArgs) Handles lblClientID.Click

        If txtClientID.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

            Me.ActiveControl = txtClientID

        Else

            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If






            Clipboard.SetText(Me.txtClientID.Text, TextDataFormat.Text)


            Me.lblClientID.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))


            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black

                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black
                lblAccountName.ForeColor = Color.Black

                lblAppName.ForeColor = Color.Black

            ElseIf lblSNoti.Visible = True Then


                lblAppName.ForeColor = Color.Blue
                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue

                lblAccountName.ForeColor = Color.Blue
                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
            End If


        End If


        lblScrapedNoti.Visible = False

        'lblPhoneRefNum.ForeColor = Color.Blue



    End Sub

    Private Sub lblOtherEmailOption_Click(sender As Object, e As EventArgs) Handles lblOtherEmailOption.Click

        If txtEmail.Text = "" Then



            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)


            Me.ActiveControl = txtOtherOptionTxt

        Else

            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If




            Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
            Clipboard.SetText(Me.txtEmail.Text, TextDataFormat.Text)

            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black

                Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblClientID.ForeColor = System.Drawing.Color.Black

                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black

                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black
                lblAccountName.ForeColor = Color.Black
                lblAppName.ForeColor = Color.Black

            ElseIf lblSNoti.Visible = True Then


                lblAppName.ForeColor = Color.Blue
                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue

                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue

                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue
            End If


            lblScrapedNoti.Visible = False


            'lblPhoneRefNum.ForeColor = Color.Blue

        End If


    End Sub

    Private Sub lblTicketTitlebar_Click(sender As Object, e As EventArgs) Handles lblTicketTitlebar.Click



        If txtScrapeBoxTitle.Text = "" Then

            MessageBox.Show("There is nothing to copy Please try again!", "Warning", MessageBoxButtons.RetryCancel)

        Else



            Me.lblTicketTitlebar.ForeColor = Color.Orange


            Clipboard.SetText(Me.txtScrapeBoxTitle.Text, TextDataFormat.Text)

            lblScrapedNoti.Visible = False

            'lblPhoneRefNum.ForeColor = Color.Blue



            If lblSNoti.Visible = True Then



                lblFirstName.ForeColor = Color.Blue
                lblLastN.ForeColor = Color.Blue
                lblPhone.ForeColor = Color.Blue
                lblEmail.ForeColor = Color.Blue
                lblEmpID.ForeColor = Color.Blue
                lblClientID.ForeColor = Color.Blue
                lblEmail.ForeColor = Color.Blue
                lblOtherEmailOption.ForeColor = Color.Blue

                lblNewHireEmail.ForeColor = Color.Blue

                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue
                lblAppName.ForeColor = Color.Blue


            Else


                lblAppName.ForeColor = Color.Black
                lblFirstName.ForeColor = Color.Black
                lblLastN.ForeColor = Color.Black
                lblPhone.ForeColor = Color.Black
                lblEmail.ForeColor = Color.Black
                lblEmpID.ForeColor = Color.Black
                lblClientID.ForeColor = Color.Black
                lblEmail.ForeColor = Color.Black
                lblOtherEmailOption.ForeColor = Color.Black

                lblNewHireEmail.ForeColor = Color.Black


                lblFirstName.ForeColor = Color.Black
                lblLastN.ForeColor = Color.Black
                lblPhone.ForeColor = Color.Black
                lblEmail.ForeColor = Color.Black
                lblEmpID.ForeColor = Color.Black
                lblClientID.ForeColor = Color.Black
                lblEmail.ForeColor = Color.Black
                lblOtherEmailOption.ForeColor = Color.Black


                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black
                lblAccountName.ForeColor = Color.Black

            End If



        End If

    End Sub

    Private Sub cboPhoneRef_SelectedIndexChanged(sender As Object, e As EventArgs)



        'If cboPhoneRef.SelectedItem = "Star Source Support" Then

        '    'lblPhoneRefNum.Text = "1-855-338-5609"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue


        'ElseIf cboPhoneRef.SelectedItem = "CVS HRSS" Then


        '    'lblPhoneRefNum.Text = "401-770-8033"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue

        'ElseIf cboPhoneRef.SelectedItem = "CVS HRSS Fax#" Then


        '    'lblPhoneRefNum.Text = "401-652-1119"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue

        'ElseIf cboPhoneRef.SelectedItem = "CVS Drug Screen / BI Support (Team West)" Then


        '    'lblPhoneRefNum.Text = "1-888-324-2103"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue

        'ElseIf cboPhoneRef.SelectedItem = "Premier" Then


        '    'lblPhoneRefNum.Text = "1-866-439-7179"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue



        'ElseIf cboPhoneRef.SelectedItem = "Missing Info Fax# (New Cases)" Then


        '    'lblPhoneRefNum.Text = "1-800-933-1875"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue

        'ElseIf cboPhoneRef.SelectedItem = "Missing Info Fax# (Authorizations)" Then


        '    'lblPhoneRefNum.Text = "1-800-213-4937"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue


        'ElseIf cboPhoneRef.SelectedItem = "Missing Info Fax# (Documentation)" Then


        '    'lblPhoneRefNum.Text = "1-888-719-8911"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue


        'ElseIf cboPhoneRef.SelectedItem = "Drug Screening" Then


        '    'lblPhoneRefNum.Text = "1-800-521-5791"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue


        'ElseIf cboPhoneRef.SelectedItem = "Consumer Disclosure" Then


        '    'lblPhoneRefNum.Text = "1-800-845-6004"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue


        'ElseIf cboPhoneRef.SelectedItem = "Lexis Nexis Help Desk" Then


        '    'lblPhoneRefNum.Text = "770-752-3201"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue


        'ElseIf cboPhoneRef.SelectedItem = "Fresenius I-9" Then


        '    'lblPhoneRefNum.Text = "1-855-362-6247"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue


        'ElseIf cboPhoneRef.SelectedItem = "Resident Screening" Then


        '    'lblPhoneRefNum.Text = "1-800-487-3240"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue


        'ElseIf cboPhoneRef.SelectedItem = "Fingerprinting (NetMark)" Then


        '    'lblPhoneRefNum.Text = "1-877-491 1752"

        '    'lblPhoneRefNum.Visible = True

        '    'lblPhoneRefNum.ForeColor = Color.Blue



        'End If





    End Sub

    Private Sub SpecificSearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpecificSearchToolStripMenuItem.Click


        SPCRefSearch.Show()
        Me.Hide()


    End Sub

    Private Sub radApplicant_CheckedChanged(sender As Object, e As EventArgs) Handles radApplicantEnt.CheckedChanged


        Me.txtEmail.TabIndex = 4

        Me.txtApplicationID.TabIndex = 5

        Me.txtOrderID.TabIndex = 6

        Me.txtAccountName.TabIndex = 7




        btnFind.Visible = False
        btnFind.Enabled = False

        Button8.Enabled = True
        Button8.Visible = True

        Button7.Enabled = False
        Button7.Visible = False


        lblAppID.Text = "Application ID:"



        Me.txtEmail.Location = New System.Drawing.Point(119, 209)
        Me.lblEmail.Location = New System.Drawing.Point(71, 216)

        Me.txtApplicationID.Location = New System.Drawing.Point(119, 236)
        Me.lblAppID.Location = New System.Drawing.Point(12, 243)


        Me.txtOrderID.Location = New System.Drawing.Point(119, 262)
        Me.lblOrderID.Location = New System.Drawing.Point(55, 269)


        Me.lblAccountName.Location = New System.Drawing.Point(9, 297)
        Me.txtAccountName.Location = New System.Drawing.Point(119, 290)

      
        lblEmail.Visible = True
        txtEmail.Visible = True

        lblOrderID.Visible = True
        txtOrderID.Visible = True

        lblAppID.Visible = True
        txtApplicationID.Visible = True
        lblAccountName.Visible = True
        txtAccountName.Visible = True


        If lblEmpID.Visible = True Then
            lblEmpID.Visible = False
        End If

        If txtUserID.Visible = True Then
            txtUserID.Visible = False
        End If


        If lblOtherEmailOption.Visible = True Then
            lblOtherEmailOption.Visible = False
        End If


        If txtOtherOptionTxt.Visible = True Then
            txtOtherOptionTxt.Visible = False

        End If

        If txtClientID.Visible = True Then
            txtClientID.Visible = False
        End If

        If lblClientID.Visible = True Then
            lblClientID.Visible = False
        End If


        If lblNewHireEmail.Visible = True Then

            lblNewHireEmail.Visible = False
        End If

        If txtNewHireEmail.Visible = True Then

            txtNewHireEmail.Visible = False
        End If

        If txtAppName.Visible = True Then

            txtAppName.Visible = False
        End If

        If lblAppName.Visible = True Then

            lblAppName.Visible = False
        End If

        ''




    End Sub

    Private Sub radClient_CheckedChanged(sender As Object, e As EventArgs) Handles radClient.CheckedChanged






        Me.txtUserID.TabIndex = 4
        Me.txtClientID.TabIndex = 5
        Me.txtOrderID.TabIndex = 6
        Me.txtAccountName.TabIndex = 7
        Me.txtEmail.TabIndex = 8



        Button8.Enabled = False
        Button8.Visible = False

        btnFind.Visible = False
        btnFind.Enabled = False

        Button7.Enabled = True
        Button7.Visible = True



        ''Put Account Name in the right field

        Me.lblAccountName.Location = New System.Drawing.Point(9, 297)
        Me.txtAccountName.Location = New System.Drawing.Point(119, 290)

        ''Put email in the right field 
        '   Me.txtOtherOptionTxt.Location = New System.Drawing.Point(106, 324)
        Me.txtEmail.Location = New System.Drawing.Point(119, 317)
        Me.lblOtherEmailOption.Location = New System.Drawing.Point(71, 324)
        Me.txtUserID.Location = New System.Drawing.Point(119, 209)
        Me.lblEmpID.Location = New System.Drawing.Point(57, 216)

        Me.txtAppName.Location = New System.Drawing.Point(119, 344)
        Me.lblAppName.Location = New System.Drawing.Point(3, 351)



        lblEmpID.Visible = True
        lblEmpID.Text = "User ID:"


        txtUserID.Visible = True
        lblOtherEmailOption.Visible = True
        ' txtOtherOptionTxt.Visible = True
        txtEmail.Visible = True

        lblOrderID.Visible = True
        txtOrderID.Visible = True
        lblClientID.Visible = True
        txtClientID.Visible = True
        '  txtNewHireEmail.Text = "NoEmail@NoEmail.com"
        lblAccountName.Visible = True
        txtAccountName.Visible = True

        txtAppName.Visible = True
        lblAppName.Visible = True




        If txtNewHireEmail.Visible = True Then

            txtNewHireEmail.Visible = False

        End If

        If lblNewHireEmail.Visible = True Then

            lblNewHireEmail.Visible = False

        End If


        If lblEmail.Visible = True Then
            lblEmail.Visible = False
        End If

        ' If txtEmail.Visible = True Then
        'txtEmail.Visible = False

        '   End If

        If txtOtherOptionTxt.Visible = True Then

            txtOtherOptionTxt.Visible = False

        End If



        If lblAppID.Visible = True Then

            lblAppID.Visible = False

        End If

        If txtApplicationID.Visible = True Then


            txtApplicationID.Visible = False

        End If




    End Sub



    Private Sub lblOrderID_Click(sender As Object, e As EventArgs) Handles lblOrderID.Click

        If txtOrderID.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)


            Me.ActiveControl = txtOrderID

        Else

            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If




            Me.lblOrderID.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                Me.lblEmpID.ForeColor = Color.Black

                Me.lblClientID.ForeColor = System.Drawing.Color.Black
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black
                lblAppName.ForeColor = Color.Black

                lblAccountName.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black

            ElseIf lblSNoti.Visible = True Then


                lblAppName.ForeColor = Color.Blue
                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = Color.Blue

                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue

                lblEmpID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue
            End If

            Clipboard.SetText(Me.txtOrderID.Text, TextDataFormat.Text)


            lblScrapedNoti.Visible = False


            'lblPhoneRefNum.ForeColor = Color.Blue

        End If





    End Sub

    Private Sub lblAppID_Click(sender As Object, e As EventArgs) Handles lblAppID.Click


        If txtApplicationID.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)


            Me.ActiveControl = txtApplicationID

        Else

            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If




            Me.lblAppID.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black
                Me.lblFirstName.ForeColor = System.Drawing.Color.Black
                Me.lblEmpID.ForeColor = Color.Black

                Me.lblClientID.ForeColor = System.Drawing.Color.Black
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black
                lblAccountName.ForeColor = Color.Black

                lblOrderID.ForeColor = Color.Black
                lblAppName.ForeColor = Color.Black


            ElseIf lblSNoti.Visible = True Then


                lblAppName.ForeColor = Color.Blue
                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                Me.lblFirstName.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = Color.Blue

                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue

                lblOrderID.ForeColor = Color.Blue
                lblAccountName.ForeColor = Color.Blue

            End If

            Clipboard.SetText(Me.txtApplicationID.Text, TextDataFormat.Text)


            lblScrapedNoti.Visible = False


            'lblPhoneRefNum.ForeColor = Color.Blue

        End If





    End Sub

    Private Sub CallCounterTimer_Tick(sender As Object, e As EventArgs) Handles CallCounterTimer.Tick


        ''Stored Notification

        lblSNoti.Visible = True



        ''Call Counter code

        CallCounter = CallCounter + 1

        lblcCounter.Text = CallCounter.ToString

        CallCounterTimer.Enabled = False









    End Sub

    Private Sub CallListTimer_Tick(sender As Object, e As EventArgs) Handles CallListTimer.Tick

        ''Show the Call List Indicator

        lblCallList_Indicator.Visible = True

        CallListTimer.Enabled = False

    End Sub

    Private Sub LogLaterCallListTimer_Tick(sender As Object, e As EventArgs) Handles LogLaterCallListTimer.Tick

        Try

            CallList()


            LogLaterCallListTimer.Enabled = False

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try


    End Sub


    Private Sub lblPhoneRefNum_Click(sender As Object, e As EventArgs)

        'Clipboard.SetText(Me.'lblPhoneRefNum.Text, TextDataFormat.Text)

        If lblTicketTitlebar.ForeColor = Color.Orange Then


            lblTicketTitlebar.ForeColor = Color.Black

        End If


        'lblPhoneRefNum.ForeColor = Color.Lime

        If lblSNoti.Visible = True Then

            Me.lblFirstName.ForeColor = Color.Blue
            Me.lblLastN.ForeColor = System.Drawing.Color.Blue
            Me.lblPhone.ForeColor = System.Drawing.Color.Blue
            Me.lblEmail.ForeColor = System.Drawing.Color.Blue
            Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue

            Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
            Me.lblClientID.ForeColor = System.Drawing.Color.Blue

            Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue


            lblAccountName.ForeColor = Color.Blue

            lblScrapedNoti.Visible = False


            lblAppID.ForeColor = Color.Blue
            lblOrderID.ForeColor = Color.Blue


        ElseIf lblFirstName.ForeColor = Color.Black Then

            Me.lblFirstName.ForeColor = Color.Black
            Me.lblLastN.ForeColor = System.Drawing.Color.Black
            Me.lblPhone.ForeColor = System.Drawing.Color.Black
            Me.lblEmail.ForeColor = System.Drawing.Color.Black
            Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black

            Me.lblEmpID.ForeColor = System.Drawing.Color.Black
            Me.lblClientID.ForeColor = System.Drawing.Color.Black

            Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black

            lblAppID.ForeColor = Color.Black
            lblOrderID.ForeColor = Color.Black
            lblAccountName.ForeColor = Color.Black
        End If


        If lblSNoti.Visible = True Then

            Me.lblFirstName.ForeColor = Color.Blue
            Me.lblLastN.ForeColor = System.Drawing.Color.Blue
            Me.lblPhone.ForeColor = System.Drawing.Color.Blue
            Me.lblEmail.ForeColor = System.Drawing.Color.Blue
            Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue

            Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
            Me.lblClientID.ForeColor = System.Drawing.Color.Blue

            Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue


            lblAppID.ForeColor = Color.Blue
            lblOrderID.ForeColor = Color.Blue
            lblAccountName.ForeColor = Color.Blue
        End If
    End Sub

    Private Sub Time_Tick(sender As Object, e As EventArgs) Handles Time.Tick


        lblNewDate.Text = Date.Now.ToString("MMM dd yyyy")
        lblNewTime.Text = CStr(TimeOfDay)



    End Sub

    Private Sub ScratchPadMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing


        If MessageBox.Show("Are you sure to close this application?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            End



        Else
            e.Cancel = True


        End If




    End Sub

    Private Sub WorkdayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WorkdayToolStripMenuItem.Click


        Process.Start(workd)


    End Sub

    Private Sub ConfluenceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConfluenceToolStripMenuItem.Click


        Process.Start(conflu)

    End Sub

    Private Sub DatabaseSearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatabaseSearchToolStripMenuItem.Click

        Try



            If DatabaseSearch.Visible = True Then

                MsgBox("The Database is already Open", 0 Or 48, "Alert")

            Else


                Dim result2 As Integer = MessageBox.Show("You are about to load the database do you want to proceed?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question)


                If result2 = DialogResult.Yes Then



                    Form1.Show()

                    DatabaseloadTimer.Enabled = True


                ElseIf result2 = DialogResult.No Then



                End If


            End If





        Catch ex As SystemException

            MsgBox(ex.Message)

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub SiebelRelaunchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SiebelRelaunchToolStripMenuItem.Click




        Siebel.WebBrowser1.Navigate("https://crm.fadv.com/fins_enu/start.swe?SWECmd=GotoView&SWEView=Contact+Screen+Homepage+View&SWERF=1&SWEHo=crm.fadv.com&SWEBU=1")


        Siebel.Show()

        Siebel.WebBrowser1.Refresh()

    End Sub

    Private Sub SignOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SignOutToolStripMenuItem.Click


        Close()



    End Sub

    Private Sub BtnSS_Click(sender As Object, e As EventArgs) Handles BtnSS.Click

        Try



            If DatabaseSearch.Visible = True Then

                MsgBox("The Database is already Open", 0 Or 48, "Alert")

            Else


                Dim result2 As Integer = MessageBox.Show("You are about to load the database do you want to proceed?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question)


                If result2 = DialogResult.Yes Then



                    Form1.Show()

                    DatabaseloadTimer.Enabled = True


                ElseIf result2 = DialogResult.No Then



                End If


            End If



        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try

    End Sub


    Private Sub btnSignIn_Click(sender As Object, e As EventArgs)

        Try


            Siebel.LogIn()

        Catch ex As NullReferenceException

            MsgBox("Something went wrong while trying to grant you access, your Seibel Password may have recently been changed if so please reach out to the dev Eric.Durrant@Fadv.com", 0 Or 48, "Alert!")

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try




    End Sub

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click

        Try


            If radManager.Checked = False And radNewHire.Checked = False And radContractor.Checked = False And radOther.Checked = False And radClient.Checked = False And radApplicantEnt.Checked = False And radFingerPrinting.Checked = False And radPROM.Checked = False Then

                MessageBox.Show("In order to use this option, please select a Job Title and try again", "Warning", MessageBoxButtons.RetryCancel)

                Me.Cursor = Cursors.Hand


            Else

                ''Send to Backender
                SendtoBackender()

                Findcontact2()





                ''Send to Backender
                SendtoBackender()



            End If


        Catch ex As NullReferenceException

            MsgBox("There was a slight issue with 'btnFind' , Please retry. If this error persist please reach out to the developer @ Eric.Durrant@Fadv.com", 0 Or 48, "Alert!")

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try


    End Sub



    Public Sub Findcontact2()

        Try



            '' If New Hire is selected

            If radNewHire.Checked = True Then


                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", "CVS")

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", "Applicant")

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", "NoEmail@NoEmail.com")

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable SR timer for New Hire

                Siebel.EnableSRNewHire2.Enabled = True

            End If

            If radManager.Checked = True Then



                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", BackEnder.txtLast.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", BackEnder.txtName.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").SetAttribute("value", BackEnder.txtPhone.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail3.Text)

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True




            End If

            If radPROM.Checked = True Then


                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", BackEnder.txtLast.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", BackEnder.txtName.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").SetAttribute("value", BackEnder.txtPhone.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail3.Text)

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True


            End If


            If radFingerPrinting.Checked = True Then


                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", "Fingerprint")

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", "Operator")


                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True

            End If



            If radContractor.Checked = True Then



                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail1.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()


                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable SR timer for New Hire
                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist5.Enabled = True



            End If

            If radOther.Checked = True Then

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", "Transfer")

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", "Applicant")


                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True


            End If




        Catch ex As NullReferenceException

            MsgBox("Make sure you are on the Contacts Tab in Siebel, this option will not work if you are on a diffrent tab", 0 Or 48, "Alert!")



        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub


    Public Sub Baller()

        Try



            '' If New Hire is selected

            If radNewHire.Checked = True Then


                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", "CVS")

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", "Applicant")

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", "NoEmail@NoEmail.com")

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable SR timer for New Hire

                Siebel.EnableSRNewHire2.Enabled = True



            ElseIf radManager.Checked = True Then

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", BackEnder.txtLast.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", BackEnder.txtName.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").SetAttribute("value", BackEnder.txtPhone.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail3.Text)

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True




            ElseIf radContractor.Enabled = True Then

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", BackEnder.txtLast.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", BackEnder.txtName.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").SetAttribute("value", BackEnder.txtPhone.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail3.Text)

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True



            ElseIf radOther.Enabled = True Then


                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", BackEnder.txtLast.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", BackEnder.txtName.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").SetAttribute("value", BackEnder.txtPhone.Text)

                '   WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail3.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail.Text)

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True




            ElseIf radFingerPrinting.Enabled = True Then

                MsgBox("!")


                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", "Fingerprint")

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", "Operator")



                'Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", BackEnder.txtLast.Text)

                'Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", BackEnder.txtName.Text)

                'Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").SetAttribute("value", BackEnder.txtPhone.Text)

                'Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail.Text)

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True



            ElseIf radClient.Enabled = True Then



                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", BackEnder.txtLast.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", BackEnder.txtName.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").SetAttribute("value", BackEnder.txtPhone.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail3.Text)


                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True





            ElseIf radPROM.Checked = True Then

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", BackEnder.txtLast.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", BackEnder.txtName.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").SetAttribute("value", BackEnder.txtPhone.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail3.Text)

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_8_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist3.Enabled = True


            End If




        Catch ex As NullReferenceException

            MsgBox("Make sure you are on the Contacts Tab in Siebel, this option will not work if you are on a diffrent tab", 0 Or 48, "Alert!")



        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub





    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged



        If ListBox1.Text = "" Then

            MessageBox.Show("This Row is empty there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

            Me.ActiveControl = ListBox1

        Else


            Clipboard.SetText(Me.ListBox1.Text, TextDataFormat.Text)



        End If


    End Sub

    Private Sub btnExsit_Click(sender As Object, e As EventArgs) Handles btnExsit.Click
        Try

            ''Click Last Name of first entry 

            Siebel.WebBrowser1.Document.GetElementById("Last Name").InvokeMember("click")



            lblExsitorNot.Text = "E"

            ''Enable SR Button
            Siebel.EnableExsistSR_4.Enabled = True


        Catch ex As NullReferenceException

            MsgBox("There was a slight issue with 'btnExsit_Click', Please retry. If this error persist please reach out to the developer @ Eric.Durrant@Fadv.com", 0 Or 48, "Alert!")

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try

    End Sub

    Private Sub btnNewContact_Click(sender As Object, e As EventArgs) Handles btnNewContact.Click

        Try

            ''Click New Button
            Siebel.WebBrowser1.Document.GetElementById("s_1_1_8_0_Ctrl").InvokeMember("click")

            ''Enable the SR button - 500 secs 
            Siebel.EnableNewSR_5.Enabled = True


            ''Scrape over contact info
            ' Siebel.ScrapeOverInfo_6.Enabled = True

            Siebel.ScrapeOverInfo_6a.Enabled = True


        Catch ex As NullReferenceException

            MsgBox("There was a slight issue with 'btnNewContact', Please retry. If this error persist please reach out to the developer @ Eric.Durrant@Fadv.com", 0 Or 48, "Alert!")

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try


    End Sub


    Private Sub LogLater_Timer_Tick(sender As Object, e As EventArgs) Handles LogLater_Timer.Tick


        ''Show the Call List Indicator

        lblCallList_Indicator.Visible = True


        LogLater_Timer.Enabled = False


    End Sub


    Private Sub btnSR_Click(sender As Object, e As EventArgs) Handles btnSR.Click

        Try



            ' '' Clicks on Service Request tab

            Dim Element As HtmlElementCollection = Siebel.WebBrowser1.Document.All



            For Each WebElement As HtmlElement In Element

                Application.DoEvents()


                Element = Siebel.WebBrowser1.Document.GetElementsByTagName("A")

                If WebElement.GetAttribute("data-tabindex") = "tabScreen9" Then

                    WebElement.InvokeMember("click")

                End If

            Next



         


            ''Enables timer to click the new button - 3 secs
            SR_Clicks_New.Enabled = True



        Catch ex As NullReferenceException

            MsgBox("There was a slight issue with 'btnSR', Please retry. If this error persist please reach out to the developer @ Eric.Durrant@Fadv.com", 0 Or 48, "Alert!")

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try




    End Sub

    Private Sub btnLogTicketButton_Click(sender As Object, e As EventArgs) Handles btnLogTicketButton.Click

        Try
            If NEWBOX.Text = "Personal Templates" And ComboBox2.Text = "CVS I-9 Templates" Then

                MessageBox.Show("You must first select a Template", "Warning", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = ComboBox1


                Me.Cursor = Cursors.Hand


            Else

                If radManager.Checked = False And radNewHire.Checked = False And radContractor.Checked = False And radOther.Checked = False And radClient.Checked = False And radApplicantEnt.Checked = False And radFingerPrinting.Checked = False And radPROM.Checked = False Then

                    MessageBox.Show("Please be advised that a Job title must be selected in order to Log this call, please select a Job Title and try again", "Warning", MessageBoxButtons.RetryCancel)

                    Me.Cursor = Cursors.Hand


                Else

                    If MessageBox.Show("You are now about to Log your ticket do you want to proceed", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then





                    Else


                        If lblScrapedNoti.Text = "Edit Template" Then


                            MessageBox.Show("You must first save the changes to your template before you log a ticket.", "Warning", MessageBoxButtons.RetryCancel)





                        Else





                            LogTicketNew()



                        End If

                    End If


                End If


            End If


        Catch ex As NullReferenceException

            '     MsgBox("Make sure you are in the SR Tab / Authenticate before you logticket, or elements are not connected", 0 Or 48, "Alert!")

            MsgBox(ex.Message, 0 Or 48, "Alert")


            MsgBox("1")

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")


            MsgBox("2")

        End Try


    End Sub


    Public Sub LogTicketNew()

        Try


      

        ''Set The SRType

            Siebel.WebBrowser1.Document.GetElementById("s_1_1_63_0").SetAttribute("value", Label5.Text)
            Siebel.WebBrowser1.Document.GetElementById("s_1_1_63_0_icon").InvokeMember("click") '


        ''Set Primary Categoru

        Siebel.WebBrowser1.Document.GetElementById("s_1_1_158_0").SetAttribute("value", Label7.Text)
        Siebel.WebBrowser1.Document.GetElementById("s_1_1_158_0_icon").InvokeMember("click")

        ''SubCategory

        Siebel.WebBrowser1.Document.GetElementById("s_1_1_160_0").SetAttribute("value", Label8.Text)
        Siebel.WebBrowser1.Document.GetElementById("s_1_1_160_0_icon").InvokeMember("click")




        If radManager.Checked = True Then




            Siebel.Authenticate_Fill_CallDetail()

            If lblExsitorNot.Text = "E" Then



                ''Seclect the Account 

                Siebel.WebBrowser1.Document.GetElementById("s_1_1_124_0").SetAttribute("value", "CVS HEALTH")


                Siebel.WebBrowser1.Document.GetElementById("s_1_1_124_0_icon").InvokeMember("click")


                Siebel.Other_Logticket_1.Enabled = True

            End If


        ElseIf radNewHire.Checked = True Then

            ''Fill In Call Detail 

            Siebel.CalldetailFill()


            '' More info - Do Work For CVS Applicant 

            Siebel.CVS_LogTicket_1.Enabled = True


        ElseIf radPROM.Checked = True Then

            Siebel.Authenticate_Fill_CallDetail_2()


            ''Selects the Account tab
            Siebel.Other_Logticket_3.Enabled = True



        ElseIf radClient.Checked = True Then


            Siebel.CalldetailFill()



            ''---- Athenticate by checking of each field 

            '' Athenticate Email
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_1_0").InvokeMember("click")


            ''Athentcate Phone
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_2_0").InvokeMember("click")

            ''Athenticate User ID
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_5_0").InvokeMember("click")


            ''Do Work For Enterprise Client 

            Siebel.Enterprise_Logticket_1.Enabled = True


        ElseIf radApplicantEnt.Checked = True Then


            Siebel.CalldetailFill()


            ''Do Work For Enterprise Applicant 

            Siebel.CVS_LogTicket_1.Enabled = True


        ElseIf radFingerPrinting.Checked = True Then


            Siebel.CalldetailFill()


            ''---- Athenticate by checking of each field 

            '' Athenticate Email
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_1_0").InvokeMember("click")


            ''Athentcate Phone
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_2_0").InvokeMember("click")

            ''Athenticate User ID
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_5_0").InvokeMember("click")



        ElseIf radContractor.Checked = True Then

            Siebel.CalldetailFill()

            ''Do Work For Contractor


            ''---- Athenticate by checking of each field 

            '' Athenticate Email
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_1_0").InvokeMember("click")


            ''Athentcate Phone
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_2_0").InvokeMember("click")

            ''Athenticate User ID
            Siebel.WebBrowser1.Document.GetElementById("s_2_1_5_0").InvokeMember("click")


            '   Siebel.CVS_Logticket_1a.Enabled = True

            Siebel.Enterprise_Logticket_1.Enabled = True

        ElseIf radOther.Checked = True Then

            Siebel.CalldetailFill()

            ''Do Work For Transfer


            Siebel.CVS_LogTicket_1.Enabled = True




        End If







        Catch ex As Exception


            MsgBox(ex.Message, 0 Or 48, "Alert")


            MsgBox("3")


        End Try





    End Sub






    Private Sub SR_Clicks_New_Tick(sender As Object, e As EventArgs) Handles SR_Clicks_New.Tick

        Try

            ''Clicks on New Button Under the Service Request Tab



            Dim Element As HtmlElementCollection = Siebel.WebBrowser1.Document.All


            If Siebel.WebBrowser1.ReadyState = WebBrowserReadyState.Complete Then


                Element = Siebel.WebBrowser1.Document.GetElementsByTagName("button")

                For Each WebElement As HtmlElement In Element
                    Application.DoEvents()

                    If WebElement.GetAttribute("title") = "Service Requests:New" Then

                        WebElement.InvokeMember("click")

                        SR_Clicks_New.Enabled = False


                    End If



                Next

            End If



            SR_Clicks_New.Enabled = False

            ''Cicks the New SR Number

            SR_Clicks_SRNumber.Enabled = True


        Catch ex As NullReferenceException

            SR_Clicks_New.Enabled = False

            MsgBox("Please make sure you are on the correct tab option to work, if this message continues to pop up hit the 'ok' button to clear", 0 Or 48, "Alert!")

            SR_Clicks_New.Enabled = False



        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

            SR_Clicks_New.Enabled = False

        End Try


    End Sub

    Private Sub SR_Clicks_SRNumber_Tick(sender As Object, e As EventArgs) Handles SR_Clicks_SRNumber.Tick

        Try


            ''Clicks on SR Number at the top of the row

            Siebel.WebBrowser1.Document.GetElementById("SR Number").InvokeMember("click")



            SR_Clicks_SRNumber.Enabled = False


            ''Enable LogCall(!) button
            btnLogTicketButton.Enabled = True


        Catch ex As NullReferenceException

            SR_Clicks_SRNumber.Enabled = False

            MsgBox("Please make sure you are on the correct tab for this option to work, if this message continues to pop up hit the 'ok' button to clear", 0 Or 48, "Alert!")


            SR_Clicks_SRNumber.Enabled = False

            SR_Clicks_SRNumber.Enabled = False

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

            SR_Clicks_SRNumber.Enabled = False

        End Try

    End Sub



    Private Sub Enable_LogTicket_Button_Tick(sender As Object, e As EventArgs) Handles Enable_LogTicket_Button.Tick

        ''Enabels LogTicket Button

        btnNewLogLater.Enabled = True



        Enable_LogTicket_Button.Enabled = False

    End Sub

    Private Sub lblAccountName_Click(sender As Object, e As EventArgs) Handles lblAccountName.Click

        If txtAccountName.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

        Else



            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If




            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black

                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblClientID.ForeColor = System.Drawing.Color.Black

                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black


                lblFirstName.ForeColor = Color.Black

                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black
                lblAppName.ForeColor = Color.Black


            ElseIf lblSNoti.Visible = True Then


                lblAppName.ForeColor = Color.Blue
                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue

                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue


                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue


                lblFirstName.ForeColor = Color.Blue

            End If

            Me.lblAccountName.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            Clipboard.SetText(Me.txtAccountName.Text, TextDataFormat.Text)

            lblScrapedNoti.Visible = False

            'lblPhoneRefNum.ForeColor = Color.Blue





        End If




    End Sub


    Private Sub BtnSSave_Click(sender As Object, e As EventArgs)

        Try

            If radContractor.Checked = True Or radManager.Checked = True Or radClient.Checked = True Or radApplicantEnt.Checked = True Or radOther.Checked = True Then


                Siebel.WebBrowser1.Document.GetElementById("s_2_1_168_0_icon").InvokeMember("click")


                Siebel.Close_Ticket_1.Enabled = True



            End If

            If radNewHire.Checked = True Then




                Siebel.WebBrowser1.Document.GetElementById("s_2_1_168_0_icon").InvokeMember("click")



                Siebel.Close_Ticket_1.Enabled = True


            End If

        Catch ex As NullReferenceException

            MsgBox("There was a slight issue with 'BtnSSave'. If this error persist please reach out to the developer @ Eric.Durrant@Fadv.com", 0 Or 48, "Alert!")

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try


    End Sub


    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        Siebel.WebBrowser1.Focus()

        Siebel.WebBrowser1.Refresh()



    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnNewLogLater.Click


        Try





            Me.Cursor = Cursors.WaitCursor

            ''Make sure job title is selected

            If radManager.Checked = False And radNewHire.Checked = False And radOther.Checked = False And radContractor.Checked = False And radApplicantEnt.Checked = False And radClient.Checked = False And radPROM.Checked = False And radFingerPrinting.Checked = False Then

                MessageBox.Show("Please be advised that a Job title must be selected in order to save this call, please select a Job Title and try again", "Warning", MessageBoxButtons.RetryCancel)

                lblLogedLaterSTORED.Visible = False

                Me.Cursor = Cursors.Hand

            Else

                ''No Duplicate Log Laters

                If lblLogedLaterSTORED.Visible = True Then

                    MessageBox.Show("Please be advised this call has already been placed in the ‘Log Later’ box", "Warning", MessageBoxButtons.RetryCancel)


                    Me.Cursor = Cursors.Hand
                Else


                    '' Make sure all required fields filled in 

                    If txtFirstName.Text = "" Then

                        MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                        Me.ActiveControl = txtFirstName

                        lblLogedLaterSTORED.Visible = False

                        Me.Cursor = Cursors.Hand

                    Else

                        '' Make sure all required fields filled in 

                        If txtLastName.Text = "" Then

                            MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                            Me.ActiveControl = txtLastName

                            Me.Cursor = Cursors.Hand

                        Else

                            '' Make sure all required fields filled in 

                            If txtPhone.Text = "" Then


                                MessageBox.Show("Please fill out all required fields before using this option", "Warning", MessageBoxButtons.RetryCancel)

                                Me.ActiveControl = txtPhone

                                Me.Cursor = Cursors.Hand

                            Else


                                '' Make sure all required fields filled in 

                                If txtCallDetail.Text = "" Then

                                    MessageBox.Show("Please fill out 'Call Detail' section before using this option", "Warning", MessageBoxButtons.RetryCancel)

                                    Me.ActiveControl = txtCallDetail


                                    lblLogedLaterSTORED.Visible = False

                                    Me.Cursor = Cursors.Hand

                                Else

                                    ''======================================== IF the Call is already stored do this==============================================================================

                                    If lblSNoti.Visible = True Then

                                        '' LogTicket Counter / pending ticket info for log later box


                                        LogLaterCouner = LogLaterCouner + 1


                                        '  Log_Later.lblPendingTicketsNumber.Text = LogLaterCouner.ToString




                                        ''Call List Indicator

                                        LogLater_Timer.Enabled = True

                                        lblLogedLaterSTORED.Visible = True



                                        ''=========================== Call information getting sent to logLater Box================================================================================


                                        Log_Later.ListBox1.Items.Add(lblNewDate.Text + " / " + lblNewTime.Text)
                                        Log_Later.ListBox1.Items.Add("Call#: " + lblcCounter.Text)
                                        Log_Later.ListBox1.Items.Add("SPC Reference #: " + lblSPCRefNum.Text)

                                        If radNewHire.Checked = True Then

                                            Log_Later.ListBox1.Items.Add("CVS New Hire")

                                        ElseIf radManager.Checked = True Then


                                            Log_Later.ListBox1.Items.Add("CVS Store Manager")

                                        ElseIf radContractor.Checked = True Then

                                            Log_Later.ListBox1.Items.Add("Field Colleague Trainer")

                                        ElseIf radOther.Checked = True Then

                                            Log_Later.ListBox1.Items.Add("Warm Transfer")



                                        ElseIf radFingerPrinting.Checked = True Then

                                            Log_Later.ListBox1.Items.Add("FingerPrinting")

                                        ElseIf radClient.Checked = True Then

                                            Log_Later.ListBox1.Items.Add("Enterprise Client")


                                        ElseIf radApplicantEnt.Checked = True Then

                                            Log_Later.ListBox1.Items.Add("Enterprise Applicant")

                                        ElseIf radPROM.Checked = True Then

                                            Log_Later.ListBox1.Items.Add("PROM")


                                        End If



                                        Log_Later.ListBox1.Items.Add(txtFirstName.Text)
                                        Log_Later.ListBox1.Items.Add(txtLastName.Text)
                                        Log_Later.ListBox1.Items.Add(txtPhone.Text)

                                        If radNewHire.Checked = True Then

                                            Log_Later.ListBox1.Items.Add(txtEmail.Text)
                                            Log_Later.ListBox1.Items.Add("1-1134812539086")
                                            Log_Later.ListBox1.Items.Add("N/a")
                                            Log_Later.ListBox1.Items.Add("N/a")
                                            Log_Later.ListBox1.Items.Add("N/a")

                                        ElseIf radManager.Checked = True Then

                                            Log_Later.ListBox1.Items.Add(txtUserID.Text)
                                            Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)
                                            Log_Later.ListBox1.Items.Add("1-1134812539086")
                                            Log_Later.ListBox1.Items.Add("N/a")
                                            Log_Later.ListBox1.Items.Add("N/a")

                                            If txtAppName.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtAppName.Text)

                                            End If

                                        ElseIf radContractor.Checked = True Then

                                            Log_Later.ListBox1.Items.Add(txtUserID.Text)
                                            Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)
                                            Log_Later.ListBox1.Items.Add("1-1134812539086")
                                            Log_Later.ListBox1.Items.Add("N/a")
                                            Log_Later.ListBox1.Items.Add("N/a")

                                            If txtAppName.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtAppName.Text)

                                            End If
                                            '' Transfer ===================================================================================================================================
                                        ElseIf radOther.Checked = True Then


                                            If txtUserID.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtUserID.Text)

                                            End If


                                            If txtClientID.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtClientID.Text)

                                            End If


                                            If txtEmail.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtEmail.Text)

                                            End If

                                            If txtAccountName.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtAccountName.Text)

                                            End If

                                            Log_Later.ListBox1.Items.Add("N/a")


                                            '' FingerPrinting ================================================================================================================================================

                                        ElseIf radFingerPrinting.Checked = True Then


                                            If txtEmail.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")
                                            Else

                                                Log_Later.ListBox1.Items.Add(txtEmail.Text)

                                            End If


                                            If txtAccountName.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else

                                                Log_Later.ListBox1.Items.Add(txtAccountName.Text)

                                            End If

                                            Log_Later.ListBox1.Items.Add("N/a")
                                            Log_Later.ListBox1.Items.Add("N/a")
                                            Log_Later.ListBox1.Items.Add("N/a")

                                            '' Enterprise Client ===============================================================================================================================================
                                        ElseIf radClient.Checked = True Then


                                            If txtUserID.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtUserID.Text)

                                            End If


                                            If txtClientID.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtClientID.Text)

                                            End If




                                            If txtOrderID.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtOrderID.Text)

                                            End If



                                            If txtEmail.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtEmail.Text)

                                            End If

                                            If txtAccountName.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtAccountName.Text)

                                            End If

                                            If txtAppName.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtAppName.Text)

                                            End If

                                            '' EnterPrise Applicant ============================================================================================================================================

                                        ElseIf radApplicantEnt.Checked = True Then



                                            If txtEmail.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtEmail.Text)

                                            End If

                                            If txtApplicationID.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtApplicationID.Text)

                                            End If



                                            If txtOrderID.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else
                                                Log_Later.ListBox1.Items.Add(txtOrderID.Text)

                                            End If


                                            Log_Later.ListBox1.Items.Add(txtAccountName.Text)
                                            Log_Later.ListBox1.Items.Add("N/a")


                                            '===========PROM==============================================================================================================================================

                                        ElseIf radPROM.Checked = True Then

                                            If txtUserID.Text = "" Then

                                                Log_Later.ListBox1.Items.Add("N/a")

                                            Else

                                                Log_Later.ListBox1.Items.Add(txtUserID.Text)

                                            End If

                                            If txtNewHireEmail.Text = "" Then


                                                Log_Later.ListBox1.Items.Add("N/a")
                                            Else

                                                Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)

                                            End If




                                            Log_Later.ListBox1.Items.Add("1-1138742798839")
                                            Log_Later.ListBox1.Items.Add("N/a")
                                            Log_Later.ListBox1.Items.Add("N/a")

                                        End If

                                        Log_Later.ListBox1.Items.Add(txtCallDetail.Text)
                                        Log_Later.ListBox1.Items.Add("!!!!!!!!!!!!!!--- Log This Call--!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                                        Log_Later.ListBox1.Items.Add("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")




                                        Log_Later.Show()

                                        Me.Cursor = Cursors.Hand


                                    Else

                                        ''If ticket is NOT stored do this==================================================================================================


                                        ''================= LogTicket Counter =============================================================================================



                                        LogLaterCouner = LogLaterCouner + 1


                                        '' pending ticket info for log later box

                                        '  Log_Later.lblPendingTicketsNumber.Text = LogLaterCouner.ToString




                                        ''Call Counter Timer ( counts call)
                                        CallCounterTimer.Enabled = True



                                        ''Store Call Thread ( stores call to database)


                                        LoglaterStoreCall = New System.Threading.Thread(AddressOf LogStoreCall)

                                        LoglaterStoreCall.Start()



                                        ''Call List Indicator

                                        LogLater_Timer.Enabled = True



                                        ''Call List Timer( sends call to call list)
                                        LogLaterCallListTimer.Enabled = True



                                        ''========== Turn Labels Blue ===========================================================================================================================================

                                        Me.lblFirstName.ForeColor = Color.Blue
                                        Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                                        Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                                        Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                                        Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue
                                        Me.lblAccountName.ForeColor = Color.Blue



                                        Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                                        Me.lblClientID.ForeColor = System.Drawing.Color.Blue

                                        Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue

                                        lblOrderID.ForeColor = Color.Blue
                                        lblAppID.ForeColor = Color.Blue

                                        lblSNoti.Visible = True

                                        Me.Cursor = Cursors.Hand

                                        lblDailyCallC.Visible = True
                                        lblcCounter.Visible = True


                                        lblReflabel.Visible = True
                                        lblSPCRefNum.Visible = True

                                        lblLogedLaterSTORED.Visible = True



                                        ''=========================== Call information getting sent to logLater Box================================================================================

                                        LogLater_CallListConnect.Enabled = True


                                        Log_Later.Show()





                                        lblLogedLaterSTORED.Visible = True



                                        Me.Cursor = Cursors.Hand



                                    End If


                                End If
                            End If

                        End If
                    End If

                End If

            End If

            ''Error Checking
        Catch ex As OleDbException

            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()




                MsgBox("The connection to the P drive was interupted..@ loglater button")


            End If

        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            MsgBox("system error at loglater button")






        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")


        End Try




    End Sub

    Private Sub LogLater_CallListConnect_Tick(sender As Object, e As EventArgs) Handles LogLater_CallListConnect.Tick



        ''=========================== Call information getting sent to logLater Box================================================================================


        Log_Later.ListBox1.Items.Add(lblNewDate.Text + " / " + lblNewTime.Text)
        Log_Later.ListBox1.Items.Add("Call#: " + lblcCounter.Text)
        Log_Later.ListBox1.Items.Add("SPC Reference #: " + lblSPCRefNum.Text)

        If radNewHire.Checked = True Then

            Log_Later.ListBox1.Items.Add("CVS New Hire")

        ElseIf radManager.Checked = True Then


            Log_Later.ListBox1.Items.Add("CVS Store Manager")

        ElseIf radContractor.Checked = True Then

            Log_Later.ListBox1.Items.Add("Field Colleague Trainer")

        ElseIf radOther.Checked = True Then

            Log_Later.ListBox1.Items.Add("Warm Transfer")



        ElseIf radFingerPrinting.Checked = True Then

            Log_Later.ListBox1.Items.Add("FingerPrinting")

        ElseIf radClient.Checked = True Then

            Log_Later.ListBox1.Items.Add("Enterprise Client")


        ElseIf radApplicantEnt.Checked = True Then

            Log_Later.ListBox1.Items.Add("Enterprise Applicant")

        ElseIf radPROM.Checked = True Then

            Log_Later.ListBox1.Items.Add("PROM")


        End If



        Log_Later.ListBox1.Items.Add(txtFirstName.Text)
        Log_Later.ListBox1.Items.Add(txtLastName.Text)
        Log_Later.ListBox1.Items.Add(txtPhone.Text)

        If radNewHire.Checked = True Then

            Log_Later.ListBox1.Items.Add(txtEmail.Text)
            Log_Later.ListBox1.Items.Add("1-1134812539086")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")

        ElseIf radManager.Checked = True Then

            Log_Later.ListBox1.Items.Add(txtUserID.Text)
            Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)
            Log_Later.ListBox1.Items.Add("1-1134812539086")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")


        ElseIf radContractor.Checked = True Then

            Log_Later.ListBox1.Items.Add(txtUserID.Text)
            Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)
            Log_Later.ListBox1.Items.Add("1-1134812539086")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")

            '' Transfer ===================================================================================================================================
        ElseIf radOther.Checked = True Then


            If txtUserID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtUserID.Text)

            End If


            If txtClientID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtClientID.Text)

            End If


            If txtEmail.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtEmail.Text)

            End If

            If txtAccountName.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtAccountName.Text)

            End If

            Log_Later.ListBox1.Items.Add("N/a")


            '' FingerPrinting ================================================================================================================================================

        ElseIf radFingerPrinting.Checked = True Then


            If txtEmail.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")
            Else

                Log_Later.ListBox1.Items.Add(txtEmail.Text)

            End If


            If txtAccountName.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else

                Log_Later.ListBox1.Items.Add(txtAccountName.Text)

            End If

            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")

            '' Enterprise Client ===============================================================================================================================================
        ElseIf radClient.Checked = True Then


            If txtUserID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtUserID.Text)

            End If


            If txtClientID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtClientID.Text)

            End If




            If txtOrderID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtOrderID.Text)

            End If



            If txtEmail.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtEmail.Text)

            End If

            Log_Later.ListBox1.Items.Add(txtAccountName.Text)

            '' EnterPrise Applicant ============================================================================================================================================

        ElseIf radApplicantEnt.Checked = True Then



            If txtEmail.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtEmail.Text)

            End If

            If txtApplicationID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtApplicationID.Text)

            End If



            If txtOrderID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else
                Log_Later.ListBox1.Items.Add(txtOrderID.Text)

            End If


            Log_Later.ListBox1.Items.Add(txtAccountName.Text)
            Log_Later.ListBox1.Items.Add("N/a")


            '===========PROM==============================================================================================================================================

        ElseIf radPROM.Checked = True Then

            If txtUserID.Text = "" Then

                Log_Later.ListBox1.Items.Add("N/a")

            Else

                Log_Later.ListBox1.Items.Add(txtUserID.Text)

            End If

            If txtNewHireEmail.Text = "" Then


                Log_Later.ListBox1.Items.Add("N/a")
            Else

                Log_Later.ListBox1.Items.Add(txtNewHireEmail.Text)

            End If




            Log_Later.ListBox1.Items.Add("1-1138742798839")
            Log_Later.ListBox1.Items.Add("N/a")
            Log_Later.ListBox1.Items.Add("N/a")

        End If

        Log_Later.ListBox1.Items.Add(txtCallDetail.Text)
        Log_Later.ListBox1.Items.Add("!!!!!!!!!!!!!!--- Log This Call--!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        Log_Later.ListBox1.Items.Add("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")


        Log_Later.Show()



        LogLater_CallListConnect.Enabled = False

        Me.Cursor = Cursors.Hand


    End Sub


    Private Sub LaucnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LaucnToolStripMenuItem.Click

        Log_Later.ShowDialog()



    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click





        Baller()







    End Sub


    Public Sub DeleteTemp()


        Try

       

        Dim con As OleDbConnection
        Dim com As OleDbCommand
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\SPC\ScratchPad Database.accdb")


        com = New OleDbCommand("delete from [Personal Templates] where [TemplateName] =@ID", con)

        con.Open()

        com.Parameters.AddWithValue("@ID", NEWBOX.Text)


            com.ExecuteNonQuery()



            RefeshCombo.Enabled = True



            MsgBox("The template has been deleted, please press 'clear box' to refresh template list")



            Me.Cursor = Cursors.Hand

        con.Close()

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try




    End Sub


    Public Sub StoreEditTemplate()


        Try


            con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\SPC\ScratchPad Database.accdb")

            con.Open()



            Dim SQL As String = "INSERT INTO [Personal Templates] ([TemplateName], [TemplateDetail], [SRType],[PrimaryCategory],[SubCategory],[User]) Values ( ?, ?, ?, ?, ?, ?)"

            Using cmd As New OleDbCommand(SQL, con)



                cmd.Parameters.AddWithValue("@p1", txtScrapeBoxTitle.Text)
                cmd.Parameters.AddWithValue("@p2", txtScrapeBox.Text)
                cmd.Parameters.AddWithValue("@p3", Label5.Text)
                cmd.Parameters.AddWithValue("@p4", Label7.Text)
                cmd.Parameters.AddWithValue("@p5", Label8.Text)
                cmd.Parameters.AddWithValue("@p6", lblAgentName.Text)


                cmd.ExecuteNonQuery()

                con.Close()



            End Using


            ''Error Checking



        Catch ex As Exception


            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try




    End Sub


    Public Sub EditTemp()

        ''Turn Scrapebox/Title Color Grey


        txtScrapeBoxTitle.BackColor = System.Drawing.SystemColors.Control

        txtScrapeBox.BackColor = System.Drawing.SystemColors.Control

        ''

        lblScrapedNoti.Text = "Edit Template"

        lblScrapedNoti.ForeColor = Color.Red

        lblScrapedNoti.Visible = True



        Me.ActiveControl = txtScrapeBox




    End Sub


    Public Sub StoreTempToData()

        Try


            NEWBOX.Items.Clear()

            FillCombo()


            ''Store Call Thread 

            TemplateStoreThread = New System.Threading.Thread(AddressOf StoreEditTemplate)

            TemplateStoreThread.Start()


            NEWBOX.Items.Clear()
            FillCombo()


        Catch ex As Exception


            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub

    Private Sub btnSaveChanges_Click(sender As Object, e As EventArgs) Handles btnSaveChanges.Click






        If lblScrapedNoti.Text <> "Edit Template" Then


            MessageBox.Show("You must first initiate the edit option before proceeding", "Warning", MessageBoxButtons.RetryCancel)


        Else

            If lblScrapedNoti.Text = "Edit Saved" Then



                MessageBox.Show("Changes have already been made, in order to make additonal changes press 'Edit Template' button", "Warning", MessageBoxButtons.RetryCancel)



            Else

                If MessageBox.Show("Do you want to save these changes?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then





                Else

                    Me.Cursor = Cursors.WaitCursor




                    MsgBox("Please wait a brief second while changes are saved")

                    ''Start Timer to start save process

                    SaveTempTimer.Enabled = True






                End If


            End If

        End If





    End Sub





    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnEditTemp.Click


        If NEWBOX.Text = "Personal Templates" Then


            MessageBox.Show("In order to Edit a template you must first select one from the 'Personal Template' drop down box", "Warning", MessageBoxButtons.RetryCancel)



        Else


            If lblScrapedNoti.Text = "Edit Saved" Then


                MessageBox.Show("Changes have already been made to this template please press the clear and retry ", "Warning", MessageBoxButtons.RetryCancel)





            Else




                If MessageBox.Show("Are you sure you want to make changes to this template?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then





                Else


                    EditTemp()




                End If


        End If



        End If











    End Sub




    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click




        btnExsit.Enabled = False

        btnNewContact.Enabled = False

        btnSR.Enabled = False

        btnLogTicketButton.Enabled = False

        ' BtnSSave.Enabled = False




    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        btnLogTicketButton.Enabled = True



    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)


        If MessageBox.Show("Are you sure the ticket is logged?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then


            Log_Ticket_Counter()


        Else

            MsgBox("Playagoodi")

        End If




    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


        ClearBackender()



    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        BackEnder.txtName.Text = txtFirstName.Text

        BackEnder.txtName.Text = txtFirstName.Text

        BackEnder.txtLast.Text = txtLastName.Text

        BackEnder.txtPhone.Text = txtPhone.Text

        BackEnder.txtEmail.Text = txtEmail.Text

        BackEnder.txtEmail1.Text = txtNewHireEmail.Text

        BackEnder.txtEmail3.Text = txtOtherOptionTxt.Text

        BackEnder.txtEmpID.Text = txtUserID.Text




        If radPROM.Checked = True Then

            BackEnder.txtVendorID.Text = txtUserID.Text


        End If

        BackEnder.txtClientId.Text = txtClientID.Text

        BackEnder.txtOrderID.Text = txtOrderID.Text

        BackEnder.txtAccountName.Text = txtAccountName.Text

        BackEnder.txtAppID.Text = txtApplicationID.Text


        BackEnder.txtCallDetail.Text = txtCallDetail.Text





    End Sub

    Private Sub TemplateCreatorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TemplateCreatorToolStripMenuItem.Click

        TemplateCreator.Show()




    End Sub




    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Try

            If radManager.Checked = False And radNewHire.Checked = False And radContractor.Checked = False And radOther.Checked = False And radClient.Checked = False And radApplicantEnt.Checked = False And radFingerPrinting.Checked = False And radPROM.Checked = False Then

                MessageBox.Show("In order to use this option, please select a Job Title and try again", "Warning", MessageBoxButtons.RetryCancel)

                Me.Cursor = Cursors.Hand


            Else

                SendtoBackender()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", BackEnder.txtEmail1.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()


                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")

                ''Enable SR timer for New Hire
                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist5.Enabled = True



            End If



        Catch ex As NullReferenceException

            MsgBox("Make sure you are on the Contacts Tab in Siebel, this option will not work if you are on a diffrent tab", 0 Or 48, "Alert!")



        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try





    End Sub


    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged


        Try

            sqltemp3 = "SELECT * FROM Templates WHERE TemplateName='" & ComboBox2.Text & " ' "



            Dim cmdtemp3 As New OleDb.OleDbCommand


            cmdtemp3.CommandText = sqltemp3
            cmdtemp3.Connection = contemp3

            readertemp3 = cmdtemp3.ExecuteReader

            If (readertemp3.Read() = True) Then

                txtScrapeBox.Text = (readertemp3("TemplateDetail"))
                txtScrapeBoxTitle.Text = (readertemp3("TemplateName"))
                Label5.Text = (readertemp3("SRType"))
                Label7.Text = (readertemp3("PrimaryCategory"))
                Label8.Text = (readertemp3("SubCategory"))



            End If

            cmdtemp3.Dispose()
            readertemp3.Close()

        Catch ex As OleDbException

            If ConnectionState.Broken = True Then



                MsgBox("The connection to the P drive was interupted..@ dropdown2")


            End If

            MsgBox("There as been an connection break, please restart for the drop down to load contents")

        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            MsgBox("system error at dropdown2")

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub

    Private Sub NEWBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NEWBOX.SelectedIndexChanged


        Try


            sqltemp2 = "SELECT * FROM [Personal Templates] WHERE TemplateName='" & NEWBOX.Text & " ' "

            Dim cmdtemp As New OleDb.OleDbCommand


            cmdtemp.CommandText = sqltemp2
            cmdtemp.Connection = contemp

            readertemp = cmdtemp.ExecuteReader

            If (readertemp.Read() = True) Then

                txtScrapeBox.Text = (readertemp("TemplateDetail"))
                txtScrapeBoxTitle.Text = (readertemp("TemplateName"))
                Label5.Text = (readertemp("SRType"))
                Label7.Text = (readertemp("PrimaryCategory"))
                Label8.Text = (readertemp("SubCategory"))



            End If

            cmdtemp.Dispose()
            readertemp.Close()


        Catch ex As OleDbException

            If ConnectionState.Broken = True Then



                MsgBox("The connection to the P drive was interupted..@ dropdown1")


            End If

            MsgBox("There as been an connection break, please restart for the drop down to load contents")


        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            MsgBox("system error at dropdown1")


        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try

            If radManager.Checked = False And radNewHire.Checked = False And radContractor.Checked = False And radOther.Checked = False And radClient.Checked = False And radApplicantEnt.Checked = False And radFingerPrinting.Checked = False And radPROM.Checked = False Then

                MessageBox.Show("In order to use this option, please select a Job Title and try again", "Warning", MessageBoxButtons.RetryCancel)

                Me.Cursor = Cursors.Hand


            Else

                SendtoBackender()



                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").SetAttribute("value", txtAccountName.Text)

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").SetAttribute("value", "Applicant")

                '  Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").SetAttribute("value", "NoEmail@NoEmail.com")

                ''Focus
                Siebel.WebBrowser1.Document.GetElementById("s_5_1_9_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_15_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_19_0").Focus()

                Siebel.WebBrowser1.Document.GetElementById("s_5_1_11_0_ctrl").InvokeMember("click")



                ''Enable SR timer for New Hire

                ''Enable New or Exsisting Button
                Siebel.EnableNew_Exsist5.Enabled = True




            End If


        Catch ex As NullReferenceException

            MsgBox("Make sure you are on the Contacts Tab in Siebel, this option will not work if you are on a diffrent tab", 0 Or 48, "Alert!")



        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try




    End Sub




    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Try


            NEWBOX.Items.Clear()

            FillCombo()


            Label10.Text = lblAgentName.Text



            Timer1.Enabled = False




        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert Timer1- main")

            Timer1.Enabled = False
        End Try

        Timer1.Enabled = False

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs)




        Me.lblAccountName.Location = New System.Drawing.Point(-1, 305)
        Me.txtAccountName.Location = New System.Drawing.Point(106, 298)

        Me.txtEmail.Location = New System.Drawing.Point(106, 217)

        lblEmail.Visible = True
        txtEmail.Visible = True

        lblOrderID.Visible = True
        txtOrderID.Visible = True

        lblAppID.Visible = True
        txtApplicationID.Visible = True
        lblAccountName.Visible = True
        txtAccountName.Visible = True


        If lblEmpID.Visible = True Then
            lblEmpID.Visible = False
        End If

        If txtUserID.Visible = True Then
            txtUserID.Visible = False
        End If


        If lblOtherEmailOption.Visible = True Then
            lblOtherEmailOption.Visible = False
        End If


        If txtOtherOptionTxt.Visible = True Then
            txtOtherOptionTxt.Visible = False

        End If

        If txtClientID.Visible = True Then
            txtClientID.Visible = False
        End If

        If lblClientID.Visible = True Then
            lblClientID.Visible = False
        End If


        If lblNewHireEmail.Visible = True Then

            lblNewHireEmail.Visible = False
        End If

        If txtNewHireEmail.Visible = True Then

            txtNewHireEmail.Visible = False
        End If

















    End Sub




    Private Sub DatabaseloadTimer_Tick(sender As Object, e As EventArgs) Handles DatabaseloadTimer.Tick





        DatabaseSearch.Show()

        Form1.Hide()

        DatabaseloadTimer.Enabled = False


    End Sub



    Private Sub lblAppName_Click(sender As Object, e As EventArgs) Handles lblAppName.Click


        If txtAppName.Text = "" Then


            MessageBox.Show("Please be advised this is an empty field, there is nothing to copy", "Warning", MessageBoxButtons.RetryCancel)

        Else



            If lblTicketTitlebar.ForeColor = Color.Orange Then


                lblTicketTitlebar.ForeColor = Color.Black

            End If




            If lblSNoti.Visible = False Then

                Me.lblLastN.ForeColor = System.Drawing.Color.Black
                Me.lblPhone.ForeColor = System.Drawing.Color.Black
                Me.lblEmail.ForeColor = System.Drawing.Color.Black
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Black

                Me.lblEmpID.ForeColor = System.Drawing.Color.Black
                Me.lblClientID.ForeColor = System.Drawing.Color.Black

                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Black


                lblFirstName.ForeColor = Color.Black

                lblOrderID.ForeColor = Color.Black
                lblAppID.ForeColor = Color.Black

                lblAccountName.ForeColor = Color.Black



            ElseIf lblSNoti.Visible = True Then


                Me.lblAppName.ForeColor = System.Drawing.Color.Blue
                Me.lblLastN.ForeColor = System.Drawing.Color.Blue
                Me.lblPhone.ForeColor = System.Drawing.Color.Blue
                Me.lblEmail.ForeColor = System.Drawing.Color.Blue
                Me.lblOtherEmailOption.ForeColor = System.Drawing.Color.Blue

                Me.lblEmpID.ForeColor = System.Drawing.Color.Blue
                Me.lblClientID.ForeColor = System.Drawing.Color.Blue
                Me.lblNewHireEmail.ForeColor = System.Drawing.Color.Blue


                lblOrderID.ForeColor = Color.Blue
                lblAppID.ForeColor = Color.Blue


                lblFirstName.ForeColor = Color.Blue

            End If

            Me.lblAppName.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))

            Clipboard.SetText(Me.txtAppName.Text, TextDataFormat.Text)

            lblScrapedNoti.Visible = False

            'lblPhoneRefNum.ForeColor = Color.Blue





        End If



    End Sub


    Private Sub DeletTempTimer_Tick(sender As Object, e As EventArgs) Handles DeletTempTimer.Tick


        Try



            Dim con As OleDbConnection
            Dim com As OleDbCommand
            con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=P:\SPC\ScratchPad Database.accdb")


            com = New OleDbCommand("delete from [Personal Templates] where [TemplateName] =@ID", con)

            con.Open()

            com.Parameters.AddWithValue("@ID", NEWBOX.Text)


            com.ExecuteNonQuery()
            '  MsgBox("Record Deleted")
            con.Close()



         



            ''Turn Scrapebox/Title Color back

            txtScrapeBoxTitle.BackColor = System.Drawing.SystemColors.ButtonHighlight

            txtScrapeBox.BackColor = System.Drawing.SystemColors.ButtonHighlight

            ''

            lblScrapedNoti.Text = "Edit Saved"

            lblScrapedNoti.ForeColor = Color.Blue


            Me.ActiveControl = txtScrapeBox


            ''
        



            DeletTempTimer.Enabled = False


            RefeshCombo.Enabled = True

       
            Me.Cursor = Cursors.Hand


                MsgBox("The template was successfully saved", 64, "Process Complete")


        Catch ex As Exception

            DeletTempTimer.Enabled = False


            MsgBox(ex.Message, 0 Or 48, "Alert")

            DeletTempTimer.Enabled = False


        End Try





    End Sub

    Public Sub RefreshCombo()

        NEWBOX.Items.Clear()

        FillCombo()


        RefeshCombo.Enabled = False



    End Sub












    Private Sub SaveTempTimer_Tick(sender As Object, e As EventArgs) Handles SaveTempTimer.Tick

        Try

     

        NEWBOX.Items.Clear()

        ''Store Call Thread 

        TemplateStoreThread = New System.Threading.Thread(AddressOf StoreEditTemplate)

        TemplateStoreThread.Start()


        SaveTempTimer.Enabled = False

        DeletTempTimer.Enabled = True





        Catch ex As Exception

            SaveTempTimer.Enabled = False


            MsgBox(ex.Message, 0 Or 48, "Alert")


            SaveTempTimer.Enabled = False


        End Try



    End Sub

    Private Sub RefeshCombo_Tick(sender As Object, e As EventArgs) Handles RefeshCombo.Tick


        Try

            NEWBOX.Items.Clear()
            FillCombo()

            RefeshCombo.Enabled = False





            RefeshCombo.Enabled = False

        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Refresh Combo")

        End Try






    End Sub

   
    Private Sub btnDeletTemp_Click(sender As Object, e As EventArgs) Handles btnDeletTemp.Click

        Try


            If NEWBOX.Text = "Personal Templates" Then


                MessageBox.Show("First select a template to delete", "Warning", MessageBoxButtons.RetryCancel)



            Else




                If lblScrapedNoti.Text = "Edit Template" Then



                    MessageBox.Show("You may not use this option while in 'Edit Template' mode", "Warning", MessageBoxButtons.RetryCancel)



                Else


                    If MessageBox.Show("Do you want to proceed with removing this template?", "Scratch Pad Compliance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then





                    Else

                        Me.Cursor = Cursors.WaitCursor



                        MsgBox("Please wait a brief second while the template is removed")


                        DeleteTemp()




                    End If


                End If

            End If




        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Delet Temp")

        End Try




    End Sub
End Class
