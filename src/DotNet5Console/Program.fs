open System
open Microsoft.Extensions.Configuration
open Authentication
open Microsoft.Graph

let loadAppSettings () =
    let appConfig = ConfigurationBuilder().AddUserSecrets<DeviceCodeAuthProvider>().Build()
    if String.IsNullOrEmpty appConfig.["appId"] || String.IsNullOrEmpty appConfig.["scopes"] then
        None
    else
        Some appConfig

let formatDateTimeTimeZone (value: DateTimeTimeZone) (dateTimeFormat: string) =
    (DateTime.Parse value.DateTime).ToString dateTimeFormat
        
let listCalendarEvents userTimeZone dateTimeFormat =
    match (GraphHelper.getCurrentWeekCalendarViewAsync DateTime.Today userTimeZone).Result with
    | None ->
        printfn "Null events."
    | Some events ->
        printfn "Events:"
        for calendarEvent in events do
            printfn $"Subject: {calendarEvent.Subject}"
            printfn $"\tOrganizer: {calendarEvent.Organizer.EmailAddress.Name}"
            printfn $"\tStart: {formatDateTimeTimeZone calendarEvent.Start dateTimeFormat}"
            printfn $"\tEnd: {formatDateTimeTimeZone calendarEvent.End dateTimeFormat}"
        
let getUserYesNo prompt =
    let rec getKeystrokesUntilYesOrNo () =
        match (Console.ReadKey true).Key with
        | ConsoleKey.Y  -> true
        | ConsoleKey.N  -> false
        | _             -> getKeystrokesUntilYesOrNo()
    
    printf "%s (y/n)" prompt
    let res = getKeystrokesUntilYesOrNo()
    printfn ""
    res
        
let rec getUserInput fieldName isRequired validate =
    printf "Enter a %s: " fieldName
    if not isRequired then
        printf "(ENTER to skip) "
    match Console.ReadLine() with
    | input when String.IsNullOrEmpty input ->
        if isRequired then
            getUserInput fieldName isRequired validate
        else
            None
    | input when validate input ->
        Some input
    | _ ->
        getUserInput fieldName isRequired validate
        
let createEvent userTimeZone =
    let rec buildAttendeeList attendees =
        match getUserInput "attendee" false (fun input -> getUserYesNo $"{input} - add attendee?") with
        | None -> attendees
        | Some newAttendee -> buildAttendeeList (newAttendee :: attendees)
    
    let subject = getUserInput "subject" true (fun input -> getUserYesNo $"Subject: {input} - is that right?")
    
    let attendeeList =
        if getUserYesNo "Do you want to invite attendees?" then
            buildAttendeeList []
        else
            []
    
    let startString = getUserInput "event start" true (fun input -> fst (DateTime.TryParse input))
    let start = DateTime.Parse startString.Value
    
    let endString = getUserInput "event end" true (fun input ->
                        match DateTime.TryParse input with
                        | true, ``end`` -> ``end`` > start
                        | _             -> false )
    let ``end`` = DateTime.Parse endString.Value
        
    let body = getUserInput "body" false (fun _ -> true)
    
    printfn $"Subject: {subject}"
    printfn $"Attendees: {String.Join(';', attendeeList)}"
    printfn $"Start: {start.ToString()}"
    printfn $"End: {``end``.ToString()}"
    printfn $"Body: {body}"
    if getUserYesNo "Create event?" then
        (GraphHelper.createEvent userTimeZone subject.Value start ``end`` attendeeList body).Wait()
    else
        printfn "Canceled."

let rec processMenu accessToken (user: User) =
    printfn "Please choose one of the following options:"
    printfn "0. Exit"
    printfn "1. Display access token"
    printfn "2. View this week calendar"
    printfn "3. Add an event"
    
    match Int32.TryParse (Console.ReadLine()) with
    | true, 0 ->
        printfn "Goodbye..."        
    | true, 1 ->
        printfn $"Access token: {accessToken}\n"
        processMenu accessToken user
    | true, 2 ->
        listCalendarEvents user.MailboxSettings.TimeZone $"{user.MailboxSettings.DateFormat} {user.MailboxSettings.TimeFormat}"
        processMenu accessToken user
    | true, 3 ->
        createEvent user.MailboxSettings.TimeZone
        processMenu accessToken user
    | _ ->
        printfn $"Invalid choice! Please try again."
        processMenu accessToken user
    
[<EntryPoint>]
let main _ =
    printfn "F# Graph Tutorial\n"
    
    match loadAppSettings() with
    | None ->
        printfn "Missing or invalid appsettings.json... Exiting."
        -1
    | Some appConfig ->
        let appId = appConfig.["appId"]
        let scopes = appConfig.["scopes"].Split(';')
        
        let authProvider = DeviceCodeAuthProvider(appId, scopes)
        
        GraphHelper.initialize authProvider
        
        match GraphHelper.getMeAsync().Result with
        | None ->
            printfn "Cannot get current user info... Exiting."
            -1
        | Some user ->
            printfn $"Welcome {user.DisplayName}!\n"
            let accessToken = authProvider.GetAccessToken().Result
            processMenu accessToken user
            0