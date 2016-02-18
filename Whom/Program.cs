using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace Whom
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials("djones9898@hotmail.com", "Fatboyz69");

            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            service.Url = new Uri(@"https://outlook.live.com/EWS/Exchange.asmx");
            //service.AutodiscoverUrl("djones9898@hotmail.com", RedirectionUrlValidationCallback);

            EmailMessage derp = (EmailMessage)service.FindItems(new FolderId(WellKnownFolderName.Inbox), new ItemView(1)).Single();
            //service.FindItems(new FolderId("INBOX"), new ItemView(1));
            //SearchMailboxesParameters searchParameters = new SearchMailboxesParameters();
            //var searchParameterQueries = (searchParameters.SearchQueries ?? new MailboxQuery[] { }).ToList() ;
            //searchParameterQueries.Add(new MailboxQuery("", new MailboxSearchScope[] { new MailboxSearchScope("inbox", MailboxSearchLocation.All) }));
            //searchParameters.SearchQueries = searchParameterQueries.ToArray();
            //var searchResults =  service.SearchMailboxes(searchParameters);


            
            /** there's alot */
            var rules = service.GetInboxRules();

            var Rule = new Rule();
            Rule.Conditions.FromAddresses.Add(derp.Sender);
            var holyshit = false;
            if (rules.Contains(Rule)) {
                holyshit = true;

            }


            if (rules.SingleOrDefault(d => d.Conditions.FromAddresses.Contains("tklarson13@hotmail.com")) != null){
                holyshit = true;

            }


            //EmailMessage email = new EmailMessage(service);

            //email.ToRecipients.Add("djones9898@hotmail.com");

            //email.Subject = "HelloWorld";
            //email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");

            //email.Send();

            var bye = Console.ReadKey();
        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
