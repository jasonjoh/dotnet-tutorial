// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.OutlookServices;

namespace dotnet_tutorial.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult SignIn()
        {
            string authority = "https://login.microsoftonline.com/common";
            string clientId = System.Configuration.ConfigurationManager.AppSettings["ida:ClientID"]; ;
            string clientSecret = System.Configuration.ConfigurationManager.AppSettings["ida:ClientSecret"]; ;
            AuthenticationContext authContext = new AuthenticationContext(authority);

            // The url in our app that Azure should redirect to after successful signin
            string redirectUri = Url.Action("Authorize", "Home", null, Request.Url.Scheme);

            // Generate the parameterized URL for Azure signin
            Uri authUri = authContext.GetAuthorizationRequestURL("https://outlook.office365.com/", clientId,
                new Uri(redirectUri), UserIdentifier.AnyUser, "prompt=login");

            // Redirect the browser to the Azure signin page
            return Redirect(authUri.ToString());
        }

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
            string redirectUri = Url.Action("Authorize", "Home", null, Request.Url.Scheme);

            // Use client ID and secret to establish app identity
            ClientCredential credential = new ClientCredential(clientId, clientSecret);

            try
            {
                // Get the token
                var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
                    authCode, new Uri(redirectUri), credential, "https://outlook.office365.com/");

                // Save the token in the session
                Session["access_token"] = authResult.AccessToken;

                return Redirect(Url.Action("Inbox", "Home", null, Request.Url.Scheme));
            }
            catch (AdalException ex)
            {
                return Content(string.Format("ERROR retrieving token: {0}", ex.Message));
            }
        }

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
                OutlookServicesClient client = new OutlookServicesClient(new Uri("https://outlook.office365.com/api/v1.0"),
                    async () =>
                    {
                        // Since we have it locally from the Session, just return it here.
                        return token;
                    });

                var mailResults = await client.Me.Messages
                                  .OrderByDescending(m => m.DateTimeReceived)
                                  .Take(10).ExecuteAsync();

                List<Models.DisplayMessage> messages = new List<Models.DisplayMessage>();
                
                foreach (Message msg in mailResults.CurrentPage)
                {
                    messages.Add(new Models.DisplayMessage(msg));
                }

                return View(messages);
            }
            catch (AdalException ex)
            {
                return Content(string.Format("ERROR retrieving messages: {0}", ex.Message));
            }
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}

// MIT License:

// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.