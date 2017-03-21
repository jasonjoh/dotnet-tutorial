// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace dotnet_tutorial.Models
{
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
            this.From = from != null ? string.Format("{0} ({1})", from.EmailAddress.Name,
                from.EmailAddress.Address) : "EMPTY";
        }
    }
}