using Microsoft.Exchange.WebServices.Data;
using System;
using System.Net;
using System.Threading.Tasks;

namespace ewsfetcher
{
    public class EWSClient
    {
        private ExchangeService _ews;
        
        public EWSClient(string url, string domain, string username, string password)
        {
            _ews = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            _ews.Credentials = new WebCredentials(username, password, domain);
            _ews.Url = new Uri(url);

            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
        }

        public Task<FindItemsResults<Item>> Fetch(TimeSpan timeLimit)
        {
            var selectedTime = DateTime.Now.Add(timeLimit);
            var searchFilter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, selectedTime);

            return _ews.FindItems(WellKnownFolderName.Inbox, searchFilter, new ItemView(50));
        }

        public Task<EmailMessage> LoadMessage(ItemId messageId)
        {
            return EmailMessage.Bind(_ews, messageId);
        }
    }
}