module GraphHelper

open System
open Microsoft.FSharp.Linq.RuntimeHelpers
open Microsoft.Graph
open FSharp.Control.Tasks
open TimeZoneConverter

let mutable graphClient : GraphServiceClient option = None
    
let initialize (authProvider: IAuthenticationProvider) =
    graphClient <- Some (GraphServiceClient(authProvider))
    
// Used to convert a F# lambda function to a LINQ expression
// Like the ones expected by the Select() method
let toLinqExpression (f: Quotations.Expr<'a>) =
    f |> LeafExpressionConverter.QuotationToLambdaExpression
    
let getMeAsync () =
    task {
        match graphClient with
        | None ->
            printfn "graphClient hasn't been initialized"
            return None
        | Some client ->
            try
                let expr =
                    <@ Func<User,obj>(fun u -> upcast {| DisplayName = u.DisplayName; MailboxSettings = u.MailboxSettings |} ) @>
                    |> toLinqExpression
                let! user = client.Me.Request().Select(expr).GetAsync()
                return Some user
            with ex ->
                printfn "Error getting signed-in user: %s" ex.Message
                return None
    }
    
let getUtcStartOfWeekInTimeZone (today: DateTime) timeZoneId =
    let userTimeZone = TZConvert.GetTimeZoneInfo timeZoneId
    let diff = System.DayOfWeek.Sunday - today.DayOfWeek
    let unspecifiedStart = DateTime.SpecifyKind(today.AddDays(float diff), DateTimeKind.Unspecified)
    TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, userTimeZone)
    
type EventProxy(subject: string, organizer: Recipient, start: DateTimeTimeZone, ``end``: DateTimeTimeZone) =
    member this.Subject = subject
    member this.Organizer = organizer
    member this.Start = start
    member this.End = ``end``

let getCurrentWeekCalendarViewAsync today timeZone =
    task {
        match graphClient with
        | None ->
            printfn "graphClient hasn't been initialized"
            return None
        | Some client ->
            let startOfWeek = getUtcStartOfWeekInTimeZone today timeZone
            let endOfWeek = startOfWeek.AddDays 7.
            
            let viewOptions : Option list = [
                QueryOption("startDateTime", startOfWeek.ToString "o") 
                QueryOption("endDateTime", endOfWeek.ToString "o")
            ]
            
            try
                // I was forced to define the EventProxy record
                // Because I kept getting errors if I created an anonymous record for the
                // Select() projection. Suggestion for getting rid of EventProxy welcomed!
                let expr =
                    <@ Func<Event,obj>(fun e -> upcast EventProxy(e.Subject, e.Organizer, e.Start, e.End) ) @>
                    |> toLinqExpression
                let! events = client.Me
                                .CalendarView
                                .Request(viewOptions)
                                .Header("Prefer", $"outlook.timezone=\"{timeZone}\"")
                                .Top(50)
                                .Select(expr)
                                .OrderBy("start/dateTime")
                                .GetAsync()
                return Some events.CurrentPage
            with :? ServiceException as ex ->
                printfn "Error getting events: %s" ex.Message
                return None
    }
    
let createEvent timeZone subject (start: DateTime) (``end``: DateTime) attendees body =
    unitTask {
        match graphClient with
        | None ->
            printfn "graphClient hasn't been initialized"
        | Some client ->
            let newEvent = Event(
                            Subject = subject,
                            Start = DateTimeTimeZone(DateTime = start.ToString "o", TimeZone = timeZone),
                            End = DateTimeTimeZone(DateTime = ``end``.ToString "o", TimeZone = timeZone))
            
            if not (List.isEmpty attendees) then
                newEvent.Attendees <- [
                    for email in attendees ->
                        Attendee(Type = AttendeeType.Required, EmailAddress = EmailAddress(Address = email))
                ]
            
            if Option.isSome body then
                newEvent.Body <- ItemBody(Content = body.Value, ContentType = BodyType.Text)
                
            try
                let! _ = client.Me.Events.Request().AddAsync(newEvent)
                printfn "Event added to calendar."
            with :? ServiceException as ex ->
                printfn "Error getting events: %s" ex.Message
    }