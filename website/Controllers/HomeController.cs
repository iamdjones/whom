using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using website.Models;

namespace website.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials("djones9898@hotmail.com", "Fatboyz69");
            service.Url = new Uri(@"https://outlook.live.com/EWS/Exchange.asmx");

            Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);




            var viewModel = new List<Sender>();

            ItemView itemView = new ItemView(10000);
            itemView.OrderBy.Add(EmailMessageSchema.DateTimeReceived, SortDirection.Descending);

            Grouping g = new Grouping();

            g.GroupOn = EmailMessageSchema.From;
            g.AggregateOn = EmailMessageSchema.DateTimeReceived;
            g.AggregateType = AggregateType.Minimum;
            g.SortDirection = SortDirection.Descending;

            SearchFilter.ContainsSubstring searchFilter = new SearchFilter.ContainsSubstring(EmailMessageSchema.From, "Nootrobox Club", ContainmentMode.Substring, ComparisonMode.IgnoreCaseAndNonSpacingCharacters);

            var groups = inbox.FindItems(itemView, g);



            Console.WriteLine("Found {0} groups", groups.Count());

            foreach (var group in groups)
            {
                var email = group.Items.First() as EmailMessage;
                //var firstItem = item.Items.First();
                var items = group.Items;
                //var groupName = (item.Items.First() as EmailMessage).Sender;
                //var groupAggregate = (item.Items.First() as EmailMessage).DateTimeReceived;

                //Console.WriteLine("{0} {1}", groupName, groupAggregate);
                var messages = group.Items;
                var sender = new Sender();
                sender.name = email.Sender.Name;
                sender.messages.AddRange(from m in messages select new Message() { subject = m.Subject, recieved = m.DateTimeReceived });

                viewModel.Add(sender);

                Console.WriteLine("{0} {1}", email.Sender, email.DateTimeReceived);
                Console.WriteLine("\t{0}", email.Subject);

            }


            return View(viewModel);
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