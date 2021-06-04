# FSMicrosoftGraph
Some Microsoft Graph -the Microsoft 365 API, samples ported to F#. As time goes by we will try to add more samples (suggestions welcomed). Right now we only have:

* Build a .NET 5 console app that connects to a Microsoft account, shows this week calendar events, and allows you to add a new event
  - This is a port of the Microsoft Graph [.NET Core Tutorial](https://docs.microsoft.com/en-us/graph/tutorials/dotnet-core)
  - Please follow the setup instructions in that tutorial, in particular:
    - Register the app in the Azure AD admin center portal, as explained in [Step 2](https://docs.microsoft.com/en-us/graph/tutorials/dotnet-core?tutorial-step=2)
    - Add Azure AD authentication to our console app, as explained in [Step 3](https://docs.microsoft.com/en-us/graph/tutorials/dotnet-core?tutorial-step=3)
    - You can also check the finished C# version [here](https://github.com/microsoftgraph/msgraph-training-dotnet-core/tree/main/demo/GraphTutorial)
    - The console app hasn't been thoroughly tested but it works with my Hotmail account 🙃

Let me know if it worked for you and what other Microsoft Graph examples you would like to see ported to F#.
