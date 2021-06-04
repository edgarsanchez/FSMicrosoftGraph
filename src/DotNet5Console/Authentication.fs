module Authentication

open System.Net.Http.Headers
open Microsoft.Graph
open Microsoft.Identity.Client
open FSharp.Control.Tasks

type DeviceCodeAuthProvider(appId, scopes) =

    let msalClient = PublicClientApplicationBuilder
                        .Create(appId)
                        .WithAuthority(AadAuthorityAudience.AzureAdAndPersonalMicrosoftAccount, true)
                        .Build()
                
    let mutable userAccount : IAccount option = None

    member _.GetAccessToken () =
        task {
            match userAccount with
            | None ->
                try
                    let! result = msalClient.AcquireTokenWithDeviceCode(scopes,
                                    fun callback -> unitTask { printfn $"{callback.Message}" } ).ExecuteAsync()
                    userAccount <- Some result.Account
                    return Some result.AccessToken
                with ex ->
                    printfn "Error getting access token: %s" ex.Message
                    return None
            | Some user ->
                let! result = msalClient.AcquireTokenSilent(scopes, user).ExecuteAsync()
                return Some result.AccessToken
        }
        
    interface IAuthenticationProvider with
        member this.AuthenticateRequestAsync request =
            unitTask {
                match! this.GetAccessToken() with
                | Some token ->
                    request.Headers.Authorization <- AuthenticationHeaderValue("bearer", token)
                | None ->
                    failwith "Couldn't get an access token."
            }