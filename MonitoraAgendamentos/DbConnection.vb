Imports System.Data.SqlClient
Imports System.Text

Imports Twilio
Imports Twilio.Rest.Api.V2010.Account
Imports Twilio.Types

Module DbConnection
    Dim cnn As String = "Data Source=LUCAS;Initial Catalog=PsicoBd;Integrated Security=True"

    Sub ConnPsi()
        ' Criação da conexão
        Using connection As New SqlConnection(cnn)
            Try
                ' Abertura da conexão
                connection.Open()

                ' Conexão estabelecida com sucesso
                Console.WriteLine("Conexão estabelecida com o banco de dados.")

                ScheduleSelect()

            Catch ex As Exception
                ' Tratamento de erros
                Console.WriteLine("Erro ao estabelecer a conexão: " & ex.Message)
            End Try
        End Using

        ' Aguarda a tecla Enter antes de fechar o console
        Console.ReadLine()
    End Sub

    Sub ScheduleSelect()
        Dim DtaAdd As String = Today.AddDays(2).ToString("yyyy-MM-dd")

        Console.WriteLine(DtaAdd)

        ' Criar a conexão com o banco de dados
        Using connection As New SqlConnection(cnn)
            connection.Open()

            ' Definir a consulta SELECT
            Dim query As String = "SELECT * FROM agenda WHERE data = '" & DtaAdd & "'"

            ' Criar o comando SQL
            Using command As New SqlCommand(query, connection)
                ' Executar o comando SELECT
                Using reader As SqlDataReader = command.ExecuteReader()
                    ' Ler os resultados da consulta
                    While reader.Read()
                        ' Acessar os valores das colunas
                        Dim id As Integer = reader.GetInt32(0)
                        Dim nome As String = reader.GetString(2)
                        Dim data As Date = reader.GetDateTime(3)
                        Dim hora As TimeSpan = reader.GetTimeSpan(4)
                        Dim Local As String = reader.GetString(5)
                        Dim celCli As String = reader.GetString(14)


                        ' Exibir os valores no console
                        Console.WriteLine($"ID: {id}, Nome: {nome}, Data: {data}, Hora: {hora}, CelCli: {celCli} , local: {Local}")

                        ' Chamar a função enviaMsg
                        EnviaMsg(celCli, nome, Format(data, "dd/MM/yyyy"), hora.ToString("hh\:mm"), Local)
                    End While
                End Using
            End Using
        End Using

        Console.ReadLine()
    End Sub

    Private Function EnviaMsg(ByVal celCli As String, ByVal nome As String, ByVal data As String, ByVal hora As String, ByVal local As String)
        Dim agora As DateTime = DateTime.Now
        Dim horaAtual As Integer = agora.Hour


        ' Substitua com suas credenciais da Twilio
        Dim accountSid As String = "ACd84971b6cd1fc3ffbe347f4a4309805c"
        Dim authToken As String = "f0a0111a6c4882be3ce7bc8261c26673"

        ' Substitua com seu número Twilio e número de telefone de destino (formato: "whatsapp:+[código do país][número de telefone]")
        Dim twilioNumber As String = "whatsapp:+14155238886"
        'Dim twilioNumber As String = "whatsapp:+5511975782527"
        Dim destinationNumber As String = "whatsapp:+55" & celCli

        ' Mensagem a ser enviada pelo WhatsApp
        Dim message As String = ""
        If horaAtual >= 6 AndAlso horaAtual < 12 Then
            ' É parte da manhã

            If local = "Presencial - Superatte" Then
                message = "Olá, *" & nome & "*! 🌞 Estamos ansiosos para a sua próxima sessão com a Psicóloga Andresa Magalhães. 📅 Lembre-se de que você tem um encontro importante agendado para o dia " & data & " às " & hora & " no *Endereço: R. Prof. João Duarte Paes, 60 - Cidade Luíza, Jundiaí*
Em caso de qualquer imprevisto, por favor, não hesite em entrar em contato conosco.
É essencial ressaltar que, para cancelar a sessão, solicitamos que nos avise com pelo menos 24 horas de antecedência. ⏰
Aproveite esse momento para se dedicar ao seu crescimento pessoal e explorar novas possibilidades. 🌱"
            Else
                message = "Olá, *" & nome & "*! 🌞 Estamos ansiosos para a sua próxima sessão com a Psicóloga Andresa Magalhães. 📅 Lembre-se de que você tem um encontro importante agendado para o dia " & data & " às " & hora & " no *Online*
Em caso de qualquer imprevisto, por favor, não hesite em entrar em contato conosco.
É essencial ressaltar que, para cancelar a sessão, solicitamos que nos avise com pelo menos 24 horas de antecedência. ⏰
Aproveite esse momento para se dedicar ao seu crescimento pessoal e explorar novas possibilidades. 🌱"
            End If

        ElseIf horaAtual >= 12 AndAlso horaAtual < 18 Then
            ' É parte da tarde
            If local = "Presencial - Superatte" Then
                message = "Olá, *" & nome & "*! 🌞 Estamos ansiosos para a sua próxima sessão com a Psicóloga Andresa Magalhães. 📅 Lembre-se de que você tem um encontro importante agendado 
para o dia *" & data & "* às *" & hora & "* no *Endereço: R. Prof. João Duarte Paes, 60 - Cidade Luíza, Jundiaí*
Em caso de qualquer imprevisto, por favor, não hesite em entrar em contato conosco.
É essencial ressaltar que, para cancelar a sessão, solicitamos que nos avise com pelo menos 24 horas de antecedência. ⏰
Aproveite esse momento para se dedicar ao seu crescimento pessoal e explorar novas possibilidades. 🌱"
            Else
                message = "Olá, *" & nome & "*! 🌞 Estamos ansiosos para a sua próxima sessão com a Psicóloga Andresa Magalhães. 📅 Lembre-se de que você tem um encontro importante agendado 
para o dia *" & data & "* às *" & hora & "* *Online*
Em caso de qualquer imprevisto, por favor, não hesite em entrar em contato conosco.
É essencial ressaltar que, para cancelar a sessão, solicitamos que nos avise com pelo menos 24 horas de antecedência. ⏰
Aproveite esse momento para se dedicar ao seu crescimento pessoal e explorar novas possibilidades. 🌱"
            End If
        Else
            If local = "Presencial - Superatte" Then
                message = "Olá, *" & nome & "*! 🌞 Estamos ansiosos para a sua próxima sessão com a Psicóloga Andresa Magalhães. 📅 Lembre-se de que você tem um encontro importante agendado 
para o dia *" & data & "* às *" & hora & "* no *Endereço: R. Prof. João Duarte Paes, 60 - Cidade Luíza, Jundiaí*
Em caso de qualquer imprevisto, por favor, não hesite em entrar em contato conosco.
É essencial ressaltar que, para cancelar a sessão, solicitamos que nos avise com pelo menos 24 horas de antecedência. ⏰
Aproveite esse momento para se dedicar ao seu crescimento pessoal e explorar novas possibilidades. 🌱"
            Else
                message = "Olá, *" & nome & "*! 🌞 Estamos ansiosos para a sua próxima sessão com a Psicóloga Andresa Magalhães. 📅 Lembre-se de que você tem um encontro importante agendado 
para o dia *" & data & "* às *" & hora & "* *Online*
Em caso de qualquer imprevisto, por favor, não hesite em entrar em contato conosco.
É essencial ressaltar que, para cancelar a sessão, solicitamos que nos avise com pelo menos 24 horas de antecedência. ⏰
Aproveite esse momento para se dedicar ao seu crescimento pessoal e explorar novas possibilidades. 🌱"
            End If
        End If

        ' Inicialize o cliente Twilio
        TwilioClient.Init(accountSid, authToken)

        ' Envie a mensagem pelo WhatsApp
        Dim messageOptions As New CreateMessageOptions(New PhoneNumber(destinationNumber))
        messageOptions.From = New PhoneNumber(twilioNumber)
        messageOptions.Body = message

        Dim sentMessage = MessageResource.Create(messageOptions)

        ' Verifique se o envio foi bem-sucedido
        If sentMessage IsNot Nothing Then
            Console.WriteLine("Mensagem enviada com sucesso para " & nome & " sessão: " & data & " e hora: " & hora)
        End If
    End Function


End Module
