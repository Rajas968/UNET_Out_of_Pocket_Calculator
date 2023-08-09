Imports Newtonsoft.Json

''' <summary>
''' Version 1.0
''' Documentation at https://data-experience.optum.com/detail/RestAPI/de.explorer%2Fdata-asset-v2%2Fapi-f5d8bcb0-19f9-4aa2-99ac-bcfc6a067e97
''' Requires apiCommonFunctions module
''' </summary>
Public Class apiAHI

#Region "Declarations"

    Public ReadOnly Property AccessToken As String
        Get
            Return tManager.FetchToken(TokenSource)
        End Get
    End Property

    Private Property Query_URI As Uri

    Private Property TokenSource As String

    Public Sub New(Optional env As Server = Server.Production)
        Select Case env
            Case Server.Production
                TokenSource = DomainProduction & TokenServer
                Query_URI = New Uri(DomainProduction & APIServerProduction & OperationQuery)
            Case Server.Release
                TokenSource = DomainTest & TokenServer
                Query_URI = New Uri(DomainProduction & APIServerRelease & OperationQuery)
            Case Server.Alpha
                TokenSource = DomainTest & TokenServer
                Query_URI = New Uri(DomainProduction & APIServerAlpha & OperationQuery)
            Case Server.Bravo
                TokenSource = DomainTest & TokenServer
                Query_URI = New Uri(DomainProduction & APIServerBravo & OperationQuery)
        End Select
    End Sub

    Enum Server
        Production
        Release
        Alpha
        Bravo
    End Enum

    'Domain
    Private Const DomainProduction As String = "https://gateway-core.optum.com"
    Private Const DomainTest As String = "https://gateway-stage-core.optum.com"

    'Token Server
    Private Const TokenServer As String = "/auth/oauth2/cached/token"

    'API Servers
    Private Const APIServerProduction As String = "/api/clm/tops-acura"
    Private Const APIServerRelease As String = "/api/rlse/clm/tops-acura"
    Private Const APIServerAlpha As String = "/api/uata/clm/tops-acura"
    Private Const APIServerBravo As String = "/api/uatb/clm/tops-acura"

    'Operation (append to server URL)
    Private Const OperationQuery As String = "/tops-members-history-claims/v1"

    'Locally-stored values for the GetNextPage function
    Private PreviousPolicy As String
    Private PreviousEmpID As String
    Private PreviousEmpname As String
    Private PreviousEmprelation As String
    Private PreviousTIN As String
    Private PreviousFirstDOS As String
    Private PreviousLastDOS As String
    Private PreviousQueryHadMoreResults As String
    Private LastICNRetrieved As String

#End Region

#Region "Functions"

    ''' <summary>
    ''' Up to 100 results. For more, use GetNextPage
    ''' </summary>
    ''' <param name="Policy">The policy number</param>
    ''' <param name="EmployeeID">10 bytes. The employee ID, preceded by a letter (typically S)</param>
    ''' <param name="PatientFirstName">Maybe limited to 10 characters? Unconfirmed</param>
    ''' <param name="RelationshipCode">2 letters</param>
    ''' <param name="TIN">All numeric. 9-digit TIN or 15-digit prefix/TIN/suffix</param>
    ''' <param name="FirstDOS">10 characters with slashes, eg 01/01/1999</param>
    ''' <param name="LastDOS">If FirstDOS was supplied, LastDOS must also be</param>
    ''' <returns>Results.Response(0).rspLineData for the collection of AHI lines. Batches of 100</returns>
    Friend Function PerformQuery(Policy As String, EmployeeID As String, PatientFirstName As String, RelationshipCode As String, Optional TIN As String = "", Optional FirstDOS As String = "", Optional LastDOS As String = "") As ahiData

        PreviousPolicy = Policy
        PreviousEmpID = EmployeeID
        PreviousEmpname = PatientFirstName
        PreviousEmprelation = RelationshipCode
        PreviousTIN = TIN
        PreviousFirstDOS = FirstDOS
        PreviousLastDOS = LastDOS

        Dim pr As New Post_Request With
            {.Request = New Post_Request.PostRequestData With
                {.reqRequiredFlds = New Post_Request.PostRequestData.RequiredFieldData With
                    {.reqSearchPolicy = Policy, .reqSearchEmpid = EmployeeID, .reqSearchEmpname = PatientFirstName, .reqSearchEmprelation = RelationshipCode, .reqSearchTIN = TIN, .reqSearchFirstDateOfService = FirstDOS, .reqSearchLastDateOfService = LastDOS, .reqNextPage = "", .reqLastICN = ""
                    }
                }
            }

        Dim jsonResult As String = sendApiRequest(Query_URI, System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(pr)), "application/json", "POST", AccessToken)

        If InStr(jsonResult, "error") > 0 Then
            Return New ahiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
        Else
            Dim tmpAhiData As New ahiData With {.Results = JsonConvert.DeserializeObject(Of Post_Response)(jsonResult), .jsonResponse = jsonResult}
            PreviousQueryHadMoreResults = tmpAhiData.Results.Response(0).stsReturnMoreDataFound
            Try
                LastICNRetrieved = tmpAhiData.Results.Response(0).rspLineData.Last.rspIcn
            Catch ex As Exception
                LastICNRetrieved = Nothing
            End Try
            Return tmpAhiData
        End If

    End Function

    ''' <summary>
    ''' For use after PerformQuery or a previous GetNextPage call.
    ''' </summary>
    ''' <returns>Returns the next batch of up to 100, or Nothing if all results have been returned.</returns>
    Friend Function GetNextPage() As ahiData

        If PreviousQueryHadMoreResults = "Y" Then
            Dim pr As New Post_Request With
            {.Request = New Post_Request.PostRequestData With
                {.reqRequiredFlds = New Post_Request.PostRequestData.RequiredFieldData With
                    {.reqSearchPolicy = PreviousPolicy, .reqSearchEmpid = PreviousEmpID, .reqSearchEmpname = PreviousEmpname, .reqSearchEmprelation = PreviousEmprelation, .reqSearchTIN = PreviousTIN, .reqSearchFirstDateOfService = PreviousFirstDOS, .reqSearchLastDateOfService = PreviousLastDOS, .reqNextPage = PreviousQueryHadMoreResults, .reqLastICN = LastICNRetrieved
                    }
                }
            }

            Dim jsonResult As String = sendApiRequest(Query_URI, System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(pr)), "application/json", "POST", AccessToken)

            If InStr(jsonResult, "error") > 0 Then
                Return New ahiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
            Else

                Dim tmpAhiData As New ahiData With {.Results = JsonConvert.DeserializeObject(Of Post_Response)(jsonResult), .jsonResponse = jsonResult}
                PreviousQueryHadMoreResults = tmpAhiData.Results.Response(0).stsReturnMoreDataFound
                Try
                    LastICNRetrieved = tmpAhiData.Results.Response(0).rspLineData.Last.rspIcn
                Catch ex As Exception
                    LastICNRetrieved = Nothing
                End Try
                Return tmpAhiData
            End If
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Will do as many queries as needed to get the full history
    ''' </summary>
    ''' <param name="Policy">The policy number</param>
    ''' <param name="EmployeeID">10 bytes. The employee ID, preceded by a letter (typically S)</param>
    ''' <param name="PatientFirstName">Maybe limited to 10 characters? Unconfirmed</param>
    ''' <param name="RelationshipCode">2 letters</param>
    ''' <param name="TIN">All numeric. 9-digit TIN or 15-digit prefix/TIN/suffix</param>
    ''' <param name="FirstDOS">10 characters with slashes, eg 01/01/1999</param>
    ''' <param name="LastDOS">If FirstDOS was supplied, LastDOS must also be</param>
    Friend Function QueryAllResults(Policy As String, EmployeeID As String, PatientFirstName As String, RelationshipCode As String, Optional TIN As String = "", Optional FirstDOS As String = "", Optional LastDOS As String = "") As ahiData

        PreviousPolicy = Policy
        PreviousEmpID = EmployeeID
        PreviousEmpname = PatientFirstName
        PreviousEmprelation = RelationshipCode
        PreviousTIN = TIN
        PreviousFirstDOS = FirstDOS
        PreviousLastDOS = LastDOS

        Dim pr As New Post_Request With
            {.Request = New Post_Request.PostRequestData With
                {.reqRequiredFlds = New Post_Request.PostRequestData.RequiredFieldData With
                    {.reqSearchPolicy = Policy, .reqSearchEmpid = EmployeeID, .reqSearchEmpname = PatientFirstName, .reqSearchEmprelation = RelationshipCode, .reqSearchTIN = TIN, .reqSearchFirstDateOfService = FirstDOS, .reqSearchLastDateOfService = LastDOS
                    }
                }
            }

        Dim jsonResult As String = sendApiRequest(Query_URI, System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(pr)), "application/json", "POST", AccessToken)

        If InStr(jsonResult, "error") > 0 Then
            Return New ahiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
        Else
            Dim FullResult As New ahiData With {.Results = JsonConvert.DeserializeObject(Of Post_Response)(jsonResult), .jsonResponse = jsonResult}
            PreviousQueryHadMoreResults = FullResult.Results.Response(0).stsReturnMoreDataFound
            Try
                LastICNRetrieved = FullResult.Results.Response(0).rspLineData.Last.rspIcn
            Catch ex As Exception
                LastICNRetrieved = Nothing
            End Try

            'Dim MoreDataFound As String = FullResult.Results.Response(0).stsReturnMoreDataFound

            Do While PreviousQueryHadMoreResults = "Y"

                Dim prNext As New Post_Request With
            {.Request = New Post_Request.PostRequestData With
                {.reqRequiredFlds = New Post_Request.PostRequestData.RequiredFieldData With
                    {.reqSearchPolicy = Policy, .reqSearchEmpid = EmployeeID, .reqSearchEmpname = PatientFirstName, .reqSearchEmprelation = RelationshipCode, .reqSearchTIN = TIN, .reqSearchFirstDateOfService = FirstDOS, .reqSearchLastDateOfService = LastDOS, .reqNextPage = "Y", .reqLastICN = FullResult.Results.Response(0).rspLineData.Last.rspIcn
                    }
                }
            }

                Dim jsonNextResult As String = sendApiRequest(Query_URI, System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(prNext)), "application/json", "POST", AccessToken)
                Dim nextgroup As New ahiData With {.Results = JsonConvert.DeserializeObject(Of Post_Response)(jsonNextResult), .jsonResponse = jsonNextResult}

                FullResult.Results.Response(0).rspLineData.AddRange(nextgroup.Results.Response(0).rspLineData)
                PreviousQueryHadMoreResults = nextgroup.Results.Response(0).stsReturnMoreDataFound
                LastICNRetrieved = nextgroup.Results.Response(0).rspLineData.Last.rspIcn
            Loop

            Return FullResult

        End If
    End Function

#End Region

#Region "Classes"

    Public Class ahiData
        Public Property Results As Post_Response
        Public jsonResponse As String
        Public IsError As Boolean = False
        Private _apiError
        Public Property apiError As apiError
            Get
                Return _apiError
            End Get
            Set(value As apiError)
                _apiError = value
                IsError = True
            End Set
        End Property
    End Class

    Public Class Post_Request
        Public Property Request As PostRequestData

        Public Class PostRequestData
            Public Property ApiConsumer As String = "NIA_Automations"
            Public Property UniqueServiceId As String = "APITOOLS"
            Public Property reqRequiredFlds As RequiredFieldData
            Public Class RequiredFieldData
                Public Property reqSearchPolicy As String = String.Empty
                Public Property reqSearchEmpid As String = String.Empty
                Public Property reqSearchEmpname As String = String.Empty
                Public Property reqSearchEmprelation As String = String.Empty
                Public Property reqSearchTIN As String = String.Empty
                Public Property reqSearchFirstDateOfService As String = String.Empty
                Public Property reqSearchLastDateOfService As String = String.Empty
                Public Property reqNextPage As String = ""
                Public Property reqLastICN As String = ""
            End Class
        End Class
    End Class

    Public Class Post_Response
        <JsonConverter(GetType(Converter(Of ResponseData)))> Property Response As New List(Of ResponseData)
    End Class

    Public Class ResponseData
        Public Property stsReturnStatus As String
        Public Property stsLoc As String
        Public Property stsCode As String
        Public Property stsCodeTypeDesc As String
        Public Property stsAddtlInfo As String
        Public Property stsReturnMoreDataFound As String
        <JsonConverter(GetType(Converter(Of RspLineDataData)))> Public Property rspLineData As New List(Of RspLineDataData)
        Public Property rspEmpLastname As String
        Public Property rspEmpAddress As String
        Public Property rspEmpCity As String
        Public Property rspEmpState As String

        Public Class RspLineDataData
            Public Property rspFirstDt As String
            Public Property rspLastDt As String
            Public Property rspTotalCharge As Double
            Public Property rspTotalCharge_String As String
            Public Property rspTotalPaid As Double
            Public Property rspTotalPaid_String As String
            Public Property rspDeductibleAmt As Double
            Public Property rspDeductibleAmt_String As String
            Public Property rspDeductibleCode As String
            Public Property rspNotCoveredAmt As Double
            Public Property rspNotCoveredAmt_String As String
            Public Property rspRemarkCode1 As String
            Public Property rspRemarkCode2 As String
            Public Property rspRemarkCode3 As String
            Public Property rspOverrideCode1 As String
            Public Property rspOverrideCode2 As String
            Public Property rspCapIndicator As String
            Public Property rspCovPhysInd As String
            Public Property rspProviderCode1 As String
            Public Property rspProviderName1 As String
            Public Property rspProviderCode2 As String
            Public Property rspProviderName2 As String
            Public Property rspIcnIndicator As String
            Public Property rspIcn As String
            Public Property rspProcDt As String
            Public Property rspDraftNumber As String

        End Class

    End Class

#End Region

End Class
