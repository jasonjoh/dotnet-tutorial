# Getting Started with the Outlook Mail API and ASP.NET #

The purpose of this guide is to walk through the process of creating a simple ASP.NET MVC C# app that retrieves messages in Office 365 or Outlook.com. The source code in this repository is what you should end up with if you follow the steps outlined here.

This tutorial will use the TODO: UPDATE LINK [Active Directory Authentication Library (Prerelease)](http://www.nuget.org/packages/Microsoft.IdentityModel.Clients.ActiveDirectory/3.5.207081303-alpha) to make OAuth2 calls and the [Microsoft Office 365 Mail, Calendar, and Contacts Library for .NET](http://www.nuget.org/packages/Microsoft.Office365.OutlookServices/) to call the Mail API.

**NOTE:** The previous version of this tutorial used the [Microsoft Office 365 API Tools](http://aka.ms/OfficeDevToolsForVS2013) to register the application in Azure AD. The registrations created with this tool are incompatible with Outlook.com, so this tutorial has been updated to use the [Application Registration Portal](https://apps.dev.microsoft.com) instead.

**NOTE:** If you are downloading this sample, you'll need to do a few things to get it to run.

1. Open the dotnet-tutorial.sln file.
2. Right-click **References** in Solution Explorer and choose **Manage NuGet Packages**.
3. Click the **Restore** button in the **Manage NuGet Packages** dialog to download all of the required packages.

## Before you begin ##

This guide assumes:

- That you already have Visual Studio 2013 installed and working on your development machine. 
- That you have an Office 365 tenant, with access to an administrator account in that tenant, **OR** an Outlook.com developer preview account.

## Create the app ##

Let's dive right in! In Visual Studio, create a new Visual C# ASP.NET Web Application. Name the application "dotnet-tutorial".

![The Visual Studio New Project window.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/new-project.PNG)

Select the `MVC` template. Click the `Change Authentication` button and choose "No Authentication". Un-select the "Host in the cloud" checkbox. The dialog should look like the following.

![The Visual Studio Template Selection window.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/template-selection.PNG)

Click OK to have Visual Studio create the project. Once that's done, run the project to make sure everything's working properly by pressing **F5** or choosing **Start Debugging** from the **Debug** menu. You should see a browser open displaying the stock ASP.NET home page. Close your browser.

Now that we've confirmed that the app is working, we're ready to do some real work.

## Designing the app ##

Our app will be very simple. When a user visits the site, they will see a button to log in and view their email. Clicking that button will take them to the Azure login page where they can login with their Office 365  or Outlook.com account and grant access to our app. Finally, they will be redirected back to our app, which will display a list of the most recent email in the user's inbox.

Let's begin by replacing the stock home page with a simpler one. Open the `./Views/Home/Index.cshtml` file. Replace the existing code with the following code.

#### Contents of the `./Views/Home/Index.cshtml` file ####

    @{
    	ViewBag.Title = "Home Page";
    }
    
    <div class="jumbotron">
	    <h1>ASP.NET MVC Tutorial</h1>
	    <p class="lead">This sample app uses the Mail API to read messages in your inbox.</p>
	    <p><a href="#" class="btn btn-primary btn-lg">Click here to login</a></p>
    </div>

This is basically repurposing the `jumbotron` element from the stock home page, and removing all of the other elements. The button doesn't do anything yet, but the home page should now look like the following.

![The sample app's home page.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/home-page.PNG)

## Implementing OAuth2 ##

Our goal in this section is to make the link on our home page initiate the [OAuth2 Authorization Code Grant flow with Azure AD](https://msdn.microsoft.com/en-us/library/azure/dn645542.aspx). To make things easier, we'll use the TODO: UPDATE LINK[Microsoft.IdentityModel.Clients.ActiveDirectory Prerelease NuGet package](http://www.nuget.org/packages/Microsoft.IdentityModel.Clients.ActiveDirectory/3.5.207081303-alpha) to handle our OAuth requests.

Before we proceed, we need to register our app to obtain a client ID and secret. Head over to https://apps.dev.microsoft.com to quickly get a client ID and secret. Using the sign in buttons, sign in with either your Microsoft account (Outlook.com), or your work or school account (Office 365).

![The Application Registration Portal Sign In Page](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/sign-in.PNG)

Once you're signed in, click the **Add an app** button. Enter `dotnet-tutorial` for the name and click **Create application**. After the app is created, locate the **Application Secrets** section, and click the **Generate New Password** button. Copy the password now and save it to a safe place. Once you've copied the password, click **Ok**.

![The new password dialog.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/new-password.PNG)

Locate the **Platforms** section, and click **Add Platform**. Choose **Web**, then enter `http://localhost:<PORT>/Home/Authorize` under **Redirect URIs**, where `<PORT>` is the port number that the Visual Studio Development Server is using for your project. You can locate this by selecting the `dotnet-tutorial` project in Solution Explorer, then checking the value for `URL` in the Properties window.

![The project properties window in Solution Explorer.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/dev-server-port.PNG)

Click **Save** to complete the registration. Copy the **Application Id** and save it along with the password you copied earlier. We'll need those values soon.

Here's what the details of your app registration should look like when you are done.

![The completed registration properties.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/dotnet-tutorial.PNG)

Open the `Web.config` file and add the following keys inside the `<appSettings>` element:

	<add key="ida:ClientID" value="YOUR APP ID" />
    <add key="ida:ClientSecret" value="YOUR APP PASSWORD" />

Replace the value of the `ida:clientID` key with the application ID you generated above, and replace the value of the `ida:ClientSecret` key with the password you generated above.

TODO: UPDATE TO INSTALL v4 The next step is to install the ADAL and Outlook libraries from NuGet. On the Visual Studio **Tools** menu, choose **NuGet Package Manager**, then **Manage NuGet Packages for Solution**. Select **Online** on the left, then enter `ADAL` in the search box in the upper-right corner. Select **Active Directory Authentication Library** from the search results and click **Install**. 

![The NuGet Package Manager window.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/nuget-package-manager.PNG)

Click through the prompts and install the package. Once the installation completes, enter `OutlookServices` in the search box, then select **Microsoft Office 365 Mail, Calendar and Contacts Library** and click **Install**.

### Back to coding ###

Now we're all set to do the sign in. Let's start by adding a `SignIn` action to the `HomeController` class. Open the `.\Controllers\HomeController.cs` file. At the top of the file, add the following lines:

	using System.Threading.Tasks;
	using Microsoft.IdentityModel.Clients.ActiveDirectory;
	using Microsoft.Office365.OutlookServices;

Add a private static string array to the `HomeController` class to hold the scopes that the app will request.

#### `scopes` array in `./Controllers/HomeController.cs` ####

	// The required scopes for our app
    private static string[] scopes = { "https://outlook.office.com/mail.read" };

Now add a new method called `SignIn` to the `HomeController` class.

#### `SignIn` action in `./Controllers/HomeController.cs` ####

	public async Task<ActionResult> SignIn()
    {
        string authority = "https://login.microsoftonline.com/common";
        string clientId = System.Configuration.ConfigurationManager.AppSettings["ida:ClientID"];
        AuthenticationContext authContext = new AuthenticationContext(authority);

        // The url in our app that Azure should redirect to after successful signin
        Uri redirectUri = new Uri("#"); // TEMPORARY

        // Generate the parameterized URL for Azure signin
        Uri authUri = await authContext.GetAuthorizationRequestUrlAsync(scopes, null, clientId, redirectUri, UserIdentifier.AnyUser, null);

        // Redirect the browser to the Azure signin page
        return Redirect(authUri.ToString());
    }

Notice that we set the `redirectUri` variable to `'#'`, which won't do a whole lot. We need to implement an action in our app that can receive a redirect back from Azure and use that URL as the value of `redirectUri`.

Add another action to the `HomeController` class called `Authorize`. This action will serve as our redirect URL.

#### `Authorize` action in `./Controllers/HomeController.cs` ####

	public ActionResult Authorize()
    {
        string authCode = Request.Params["code"];
        return Content("Auth Code: " + authCode);
    }

This doesn't do anything but display the authorization code returned by Azure, but it will let us test that we can successfully log in. Update the `SignIn` action to use the URL for the `Authorize` action for `redirectUri`.

#### Updated `SignIn` action in `./Controllers/HomeController.cs` ####

	public ActionResult SignIn()
    {
        string authority = "https://login.microsoftonline.com/common";
        string clientId = System.Configuration.ConfigurationManager.AppSettings["ida:ClientID"]; 
        AuthenticationContext authContext = new AuthenticationContext(authority);

        // The url in our app that Azure should redirect to after successful signin
        Uri redirectUri = new Uri(Url.Action("Authorize", "Home", null, Request.Url.Scheme));

        // Generate the parameterized URL for Azure signin
        Uri authUri = await authContext.GetAuthorizationRequestUrlAsync(scopes, null, clientId, redirectUri, UserIdentifier.AnyUser, null);

        // Redirect the browser to the Azure signin page
        return Redirect(authUri.ToString());
    }

Finally, let's update the home page so that the login button invokes the `SignIn` action.

#### Updated contents of the `./Views/Home/Index.cshtml` file ####

    @{
    	ViewBag.Title = "Home Page";
    }
    
    <div class="jumbotron">
	    <h1>ASP.NET MVC Tutorial</h1>
	    <p class="lead">This sample app uses the Mail API to read messages in your inbox.</p>
	    <p><a href="@Url.Action("SignIn", "Home", null, Request.Url.Scheme)" class="btn btn-primary btn-lg">Click here to login</a></p>
    </div>

Save your work and run the app. Click on the button to sign in. After signing in, you should be returned to your app, which displays an authorization code. Now let's do something with it.

### Exchanging the code for a token ###

Now let's update the `Authorize` function to retrieve a token. Replace the current code with the following.

#### Updated `Authorize` action in `./Controllers/HomeController.cs` ####

	// Note the function signature is changed!
    public async Task<ActionResult> Authorize()
    {
        // Get the 'code' parameter from the Azure redirect
        string authCode = Request.Params["code"];

        string authority = "https://login.microsoftonline.com/common";
        string clientId = System.Configuration.ConfigurationManager.AppSettings["ida:ClientID"]; ;
        string clientSecret = System.Configuration.ConfigurationManager.AppSettings["ida:ClientSecret"]; ;
        AuthenticationContext authContext = new AuthenticationContext(authority);

        // The same url we specified in the auth code request
        Uri redirectUri = new Uri(Url.Action("Authorize", "Home", null, Request.Url.Scheme));

        // Use client ID and secret to establish app identity
        ClientCredential credential = new ClientCredential(clientId, clientSecret);

        try
        {
            // Get the token
            var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
                authCode, redirectUri, credential, scopes);

			// Save the token in the session
            Session["access_token"] = authResult.Token;

            return Content("Access Token: " + authResult.Token);
        }
        catch (AdalException ex)
        {
            return Content(string.Format("ERROR retrieving token: {0}", ex.Message));
        }
    }

Save your changes and restart the app. This time, after you sign in, you should see an access token displayed. Now that we can retrieve the access token, we're ready to call the Mail API.

## Using the Mail API ##

Let's start by adding a new action to the `HomeController` class. Open the `.\Controllers\HomeController.cs` file. Add a new function to the `HomeController` class called `Inbox`.

#### `Inbox` action in `./Controllers/HomeController.cs` ####

	public async Task<ActionResult> Inbox()
    {
        string token = (string)Session["access_token"];
        if (string.IsNullOrEmpty(token))
        {
            // If there's no token in the session, redirect to Home
            return Redirect("/");
        }

        return Content(string.Format("Found token in session: {0}", token));
    }

Now update the `Authorize` function to redirect to the `Inbox` action after successfully retrieving a token.

#### Updated `Authorize` action in `./Controllers/HomeController.cs` ####

	public async Task<ActionResult> Authorize()
    {
        // Get the 'code' parameter from the Azure redirect
        string authCode = Request.Params["code"];

        string authority = "https://login.microsoftonline.com/common";
        string clientId = System.Configuration.ConfigurationManager.AppSettings["ida:ClientID"]; ;
        string clientSecret = System.Configuration.ConfigurationManager.AppSettings["ida:ClientSecret"]; ;
        AuthenticationContext authContext = new AuthenticationContext(authority);

        // The same url we specified in the auth code request
        Uri redirectUri = new Uri(Url.Action("Authorize", "Home", null, Request.Url.Scheme));

        // Use client ID and secret to establish app identity
        ClientCredential credential = new ClientCredential(clientId, clientSecret);

        try
        {
            // Get the token
            var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
                    authCode, redirectUri, credential, scopes);

            // Save the token in the session
            Session["access_token"] = authResult.Token;

            return Redirect(Url.Action("Inbox", "Home", null, Request.Url.Scheme));
        }
        catch (AdalException ex)
        {
            return Content(string.Format("ERROR retrieving token: {0}", ex.Message));
        }
    }

If you run the app now, you won't see much of a change. It still just displays the access token. The difference is we're displaying it from the Inbox action, which verifies that we succeeded in passing the token via the `Session`. Let's put it to use.

Update the `Inbox` function with the following code.

#### Updated `Inbox` action in `./Controllers/HomeController.cs` ####

	public async Task<ActionResult> Inbox()
    {
        string token = (string)Session["access_token"];
        if (string.IsNullOrEmpty(token))
        {
            // If there's no token in the session, redirect to Home
            return Redirect("/");
        }

        try
        {
            OutlookServicesClient client = new OutlookServicesClient(new Uri("https://outlook.office.com/api/v1.0"),
                async () =>
                {
                    // Since we have it locally from the Session, just return it here.
                    return token;
                });

            var mailResults = await client.Me.Messages
                              .OrderByDescending(m => m.DateTimeReceived)
							  .Take(10)
							  .Select(m => new { m.Subject, m.DateTimeReceived, m.From })
                              .ExecuteAsync();

            string content = "";

            foreach (var msg in mailResults.CurrentPage)
            {
                content += string.Format("Subject: {0}<br/>", msg.Subject);
            }

            return Content(content);
        }
        catch (AdalException ex)
        {
            return Content(string.Format("ERROR retrieving messages: {0}", ex.Message));
        }
    }

To summarize the new code in the `mail` function:

- It creates an `OutlookServicesClient` object.
- It issues a GET request to the URL for inbox messages, with the following characteristics:
	- It uses the `OrderBy()` function with a value of `DateTimeReceived desc` to sort the results by DateTimeReceived.
	- It uses the `Take()` function with a value of `10` to limit the results to the first 10.
	- It uses the `Select()` function to limit the fields returned to only those that we need.
- It loops over the results and prints out the subject.

If you restart the app now, you should get a very basic listing of email subjects. However, we can use the features of MVC to do better than that.

### Displaying the results ###

MVC can generate views based on a model. So let's start by creating a model that represents the properties of a message that we'd like to display. For our purposes, we'll choose `Subject`, `DateTimeReceived`, and `From`. In Visual Studio's Solution Explorer, right-click the `./Models` folder and choose **Add**, then **Class**. Name the class `DisplayMessage` and click **Add**.

Open the `./Models/DisplayMessage.cs` file and replace the empty class definition with the following.

#### `DisplayMessage` class definition ####

	public class DisplayMessage
    {
        public string Subject { get; set; }
        public DateTimeOffset DateTimeReceived { get; set; }
        public string From { get; set; }

        public DisplayMessage(string subject, DateTimeOffset? dateTimeReceived, 
            Microsoft.Office365.OutlookServices.Recipient from)
        {
            this.Subject = subject;
            this.DateTimeReceived = (DateTimeOffset)dateTimeReceived;
            this.From = string.Format("{0} ({1})", from.EmailAddress.Name,
                            from.EmailAddress.Address);
        }
    }

All this class does is expose the three properties of the message we want to display.

Now that we have a model, let's create a view based on it. In Solution Explorer, right-click the `./Views/Home` folder and choose **Add**, then **View**. Enter `Inbox` for the **View name**. Change the **Template** field to `List`, and choose `DisplayMessage (dotnet_tutorial.Models)` for the **Model class**. Leave everything else as default values and click **Add**.

![The Add View dialog.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/add-view.PNG)

Just one more thing to do. Let's update the `Inbox` function to use our new model and view. 

#### Updated `Inbox` action in `./Controllers/HomeController.cs` ####

	public async Task<ActionResult> Inbox()
    {
        string token = (string)Session["access_token"];
        if (string.IsNullOrEmpty(token))
        {
            // If there's no token in the session, redirect to Home
            return Redirect("/");
        }

        try
        {
            OutlookServicesClient client = new OutlookServicesClient(new Uri("https://outlook.office.com/api/v1.0"),
                async () =>
                {
                    // Since we have it locally from the Session, just return it here.
                    return token;
                });

            var mailResults = await client.Me.Messages
                              .OrderByDescending(m => m.DateTimeReceived)
                              .Take(10)
                              .Select(m => new Models.DisplayMessage(m.Subject, m.DateTimeReceived, m.From))
                              .ExecuteAsync();

            return View(mailResults.CurrentPage);
        }
        catch (AdalException ex)
        {
            return Content(string.Format("ERROR retrieving messages: {0}", ex.Message));
        }
    }

The changes here are minimal. Instead of building a string with the results, we instead create a new `DisplayMessage` object within the `Select` function. This causes the `mailResults.CurrentPage` collection to be a collection of `DisplayMessage` objects, which is perfect for our view.

Save your changes and run the app. You should now get a list of messages that looks something like this.

![The sample app displaying a user's inbox.](https://raw.githubusercontent.com/jasonjoh/dotnet-tutorial/master/readme-images/inbox-display.PNG)

## Next Steps ##

Now that you've created a working sample, you may want to learn more about the [capabilities of the Mail API](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations). If your sample isn't working, and you want to compare, you can download the end result of this tutorial from [GitHub](https://github.com/jasonjoh/dotnet-tutorial). If you download the project from GitHub, be sure to re-run the Add Connected Service wizard to register the app before trying it.

## Copyright ##

Copyright (c) Microsoft. All rights reserved.

----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Exchange Dev Blog](http://blogs.msdn.com/b/exchangedev/)