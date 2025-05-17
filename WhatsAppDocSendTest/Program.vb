Imports System
Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Module WhatsAppAPI
    Private ReadOnly httpClient As New HttpClient()

    Sub Main()
        MainAsync().GetAwaiter().GetResult()
        Console.ReadLine()
    End Sub

    Async Function MainAsync() As Task
        Await SendWhatsAppTemplate()
    End Function

    Async Function SendWhatsAppTemplate() As Task
        Dim apiKey As String = ""
        Dim phoneNumber As String = "919833533311"
        Dim recipientNumber As String = "917887678070"'Enter Recipients Number here
        Dim filePath As String = "C:/File/pATH.pdf" ' Put your local file path here
        Dim client_name As String = "Bob"
        Dim inv_amount As String = "Rs. 1000"
        
        
        Try
            ' Step 1: Upload the file and get the URL
            Dim fileUrl As String = Await UploadFile(apiKey, phoneNumber, filePath).ConfigureAwait(False)

            If String.IsNullOrEmpty(fileUrl) Then
                Console.WriteLine("Failed to upload file")
                Return
            End If

            ' Step 2: Send the template message
            Dim success As Boolean = Await SendTemplateMessage(
                apiKey, 
                phoneNumber, 
                recipientNumber, 
                fileUrl, 
                Path.GetExtension(filePath),
                client_name,
                inv_amount).ConfigureAwait(False)

            Console.WriteLine(If(success, "Message sent successfully", "Failed to send message"))

        Catch ex As Exception
            Console.WriteLine($"Error: {ex.Message}")
        End Try
    End Function

    Async Function UploadFile(apiKey As String, phoneNumber As String, filePath As String) As Task(Of String)
        If Not File.Exists(filePath) Then
            Console.WriteLine("File not found")
            Return String.Empty
        End If

        Using formData As New MultipartFormDataContent()
            httpClient.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)

            ' Read file and detect content type
            Dim fileBytes As Byte() = File.ReadAllBytes(filePath)
            Dim fileContent As New ByteArrayContent(fileBytes)
            
            ' Set proper content type based on file extension
            Dim contentType As String = GetMimeType(Path.GetExtension(filePath))
            fileContent.Headers.ContentType = New MediaTypeHeaderValue(contentType)

            formData.Add(fileContent, "file", Path.GetFileName(filePath))

            Try
                Dim response = Await httpClient.PostAsync(
                    $"https://cloudapi.wbbox.in/api/v1.0/uploads/{phoneNumber}", 
                    formData).ConfigureAwait(False)

                If response.IsSuccessStatusCode Then
                    Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                    Dim json As JObject = JObject.Parse(responseContent)
                    Console.WriteLine($"Response: {json}")
                    'Console.WriteLine($"ImageUrl: {json("ImageUrl")}")
                    Dim data = json.Value(Of JObject)("data")
                    If data IsNot Nothing Then
                        Dim imageUrl = data.Value(Of String)("ImageUrl")
                        If Not String.IsNullOrEmpty(imageUrl) Then
                            Console.WriteLine($"Successfully obtained URL: {imageUrl}")
                            Return imageUrl
                         End If
                    End If

                    'Return If(json.Value(Of String)("ImageUrl"), String.Empty)
                Else
                    Dim errorContent = Await response.Content.ReadAsStringAsync()
                    Console.WriteLine($"Upload failed ({response.StatusCode}): {errorContent}")
                    Return String.Empty
                End If
            Catch ex As Exception
                Console.WriteLine($"Upload error: {ex.Message}")
                Return String.Empty
            End Try
        End Using
    End Function

    Async Function SendTemplateMessage(
        apiKey As String, 
        phoneNumber As String, 
        recipientNumber As String, 
        fileUrl As String,
        fileExtension As String,
        client_name As String,
        inv_amount As String
        ) As Task(Of Boolean)
        
        
        httpClient.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
        httpClient.DefaultRequestHeaders.Accept.Clear()
        httpClient.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

        ' Determine the parameter type based on file extension
        Dim mediaType As String = If(
            fileExtension.Equals(".pdf", StringComparison.OrdinalIgnoreCase),
            "document",
            "image")


        ' Create components array
    Dim components As New JArray()
    
    ' Add header component
    Dim headerParameters As New JArray()
    headerParameters.Add(New JObject(
        New JProperty("type", mediaType),
        New JProperty(mediaType, New JObject(
            New JProperty("link", fileUrl),
            New JProperty("caption", "test"),
            New JProperty("filename", Path.GetFileName(fileUrl))
        ))
    ))
    
    
    components.Add(New JObject(
        New JProperty("type", "header"),
        New JProperty("parameters", headerParameters)
    ))
    
    ' Add body component
    Dim bodyParameters As New JArray()
    bodyParameters.Add(New JObject(
        New JProperty("type", "text"),
        New JProperty("text", client_name)'client_name: Anish
    ))
    bodyParameters.Add(New JObject(
        New JProperty("type", "text"),
        New JProperty("text", inv_amount)'Invoice Amount:91
    ))
    
    components.Add(New JObject(
        New JProperty("type", "body"),
        New JProperty("parameters", bodyParameters)
    ))
    
    ' Add button component
    Dim buttonParameters As New JArray()
    buttonParameters.Add(New JObject(
        New JProperty("type", "text"),
        New JProperty("text", "http://www.google.com/")
    ))
    
    components.Add(New JObject(
        New JProperty("type", "button"),
        New JProperty("parameters", buttonParameters),
        New JProperty("sub_type", "url"),
        New JProperty("index", 0)
    ))
    
    ' Build the complete template request
    Dim templateRequest As New JObject(
        New JProperty("messaging_product", "whatsapp"),
        New JProperty("recipient_type", "individual"),
        New JProperty("to", recipientNumber),
        New JProperty("type", "template"),
        New JProperty("template", New JObject(
            New JProperty("name", "test10"),
            New JProperty("language", New JObject(
                New JProperty("code", "en")
            )),
            New JProperty("components", components)
        ))
    )
        Console.WriteLine($"Template Request: {templateRequest}")
        Try
            Dim response = Await httpClient.PostAsync(
                $"https://cloudapi.wbbox.in/api/v1.0/messages/send-template/{phoneNumber}",
                New StringContent(templateRequest.ToString(), System.Text.Encoding.UTF8, "application/json")
            ).ConfigureAwait(False)

            If Not response.IsSuccessStatusCode Then
                Dim errorContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine($"Send failed ({response.StatusCode}): {errorContent}")
            End If

            Return response.IsSuccessStatusCode
        Catch ex As Exception
            Console.WriteLine($"Send error: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function GetMimeType(fileExtension As String) As String
        Select Case fileExtension.ToLower()
            Case ".jpg", ".jpeg"
                Return "image/jpeg"
            Case ".png"
                Return "image/png"
            Case ".gif"
                Return "image/gif"
            Case ".pdf"
                Return "application/pdf"
            Case ".doc"
                Return "application/msword"
            Case ".docx"
                Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            Case ".xls"
                Return "application/vnd.ms-excel"
            Case ".xlsx"
                Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Case ".ppt"
                Return "application/vnd.ms-powerpoint"
            Case ".pptx"
                Return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            Case Else
                Return "application/octet-stream"
        End Select
    End Function
End Module
