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

            //service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;
            service.Url = new Uri(@"https://outlook.live.com/EWS/Exchange.asmx");
            
            Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);
            //EmailMessage firstItem = inbox.FindItems(new ItemView(1)).First() as EmailMessage;

            


            //ResetRules(service);

            //ResetInbox(inbox);
            OrganizeInbox(service, inbox);
            
            Console.WriteLine("welp, see ya!");
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

        private static void ResetRules(ExchangeService service)
        {
            var rules = service.GetInboxRules();
            List<RuleOperation> ops = new List<RuleOperation>();
            foreach (var rule in rules)
            {

                Console.WriteLine("marking {0} for removal", rule.DisplayName);
                DeleteRuleOperation op = new DeleteRuleOperation(rule.Id);
                ops.Add(op);
            }

            if (ops.Count > 0)
            {
                Console.WriteLine("removing all marked rules");
                service.UpdateInboxRules(ops, true);
            }
            else
            {
                Console.WriteLine("no rules found");
            }

        }

        private static void ResetInbox(Folder parent)
        {

            FolderView folderView = new FolderView(10000);
            folderView.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            folderView.PropertySet.Add(FolderSchema.DisplayName);
            folderView.PropertySet.Add(FolderSchema.TotalCount);
            folderView.Traversal = FolderTraversal.Shallow;




            foreach (var folder in parent.FindFolders(folderView))
            {

                if (folder.FindFolders(folderView).Count() > 0)
                {
                    // recursion mofo
                    Console.WriteLine("Found sub folders in {0}, resetting", folder.DisplayName);
                    ResetInbox(folder);
                }

                ItemView itemView = new ItemView(10000);
                itemView.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                itemView.PropertySet.Add(EmailMessageSchema.From);
                itemView.PropertySet.Add(EmailMessageSchema.Sender);

                var folderItems = folder.FindItems(itemView);

                Console.WriteLine("Moving {0} items from {1} to {2}", folderItems.TotalCount, folder.DisplayName, parent.DisplayName);

                foreach (var item in folder.FindItems(itemView))
                {
                    item.Move(parent.Id);
                }

                Console.WriteLine("removing {0}", folder.DisplayName);
                folder.Delete(DeleteMode.HardDelete);



            }

        }

        private static void OrganizeInbox(ExchangeService service, Folder inbox)
        {

            ItemView itemView = new ItemView(10000);
            itemView.OrderBy.Add(EmailMessageSchema.DateTimeReceived, SortDirection.Descending);
            
            Grouping g = new Grouping();
            
            g.GroupOn = EmailMessageSchema.From; 
            g.AggregateOn = EmailMessageSchema.DateTimeReceived;
            g.AggregateType = AggregateType.Minimum;
            g.SortDirection = SortDirection.Descending;

            SearchFilter.ContainsSubstring searchFilter = new SearchFilter.ContainsSubstring(EmailMessageSchema.From, "Nootrobox Club", ContainmentMode.Substring, ComparisonMode.IgnoreCaseAndNonSpacingCharacters);
            
            var groups = inbox.FindItems( itemView,g);
            
            

            Console.WriteLine("Found {0} groups", groups.Count());

            foreach (var group in groups)
            {
                var email = group.Items.First() as EmailMessage;
                //var firstItem = item.Items.First();
                var items = group.Items;
                //var groupName = (item.Items.First() as EmailMessage).Sender;
                //var groupAggregate = (item.Items.First() as EmailMessage).DateTimeReceived;

                //Console.WriteLine("{0} {1}", groupName, groupAggregate);
                                               
                Console.WriteLine("{0} {1}", email.Sender, email.DateTimeReceived);
                Console.WriteLine("\t{0}", email.Subject);

                //SearchFilter folderFilter = new SearchFilter.ContainsSubstring();

                //Folder senderFolder = GetSenderFolder(service, inbox, email.Sender.Name);

                //var associatedAddress = from EmailMessage item in items
                //                        group item.Sender by item.Sender.Address into g2
                //                        select g2.FirstOrDefault();

                //foreach (var addy in associatedAddress)
                //{
                    
                //    Rule senderRule = GetSenderRule(service, addy, senderFolder);
                //    Console.WriteLine("\t{0} ({1})", addy, items.Count(i=>(i as EmailMessage).Sender.Address == addy.Address));
                //}


                //var itemIds = from EmailMessage item in items
                //              select item.Id;


                //service.MoveItems(itemIds, senderFolder.Id);

            }
        }

        public static Rule GetSenderRule(ExchangeService service, EmailAddress sender, Folder senderFolder)
        {
            Rule rule = null;
            var rules = service.GetInboxRules();
            rule = rules.SingleOrDefault(r => r.DisplayName == sender.Name);
            RuleOperation op;
            if (rule == null)
            {
                rule = new Rule();
                rule.DisplayName = sender.Name;
                op = new CreateRuleOperation(rule);
            } else
            {
                op = new SetRuleOperation(rule);
            }

            if (!rule.Conditions.ContainsSenderStrings.Contains(sender.Name))
            {
                rule.Conditions.ContainsSenderStrings.Add(sender.Name);
            }

            if (!rule.Conditions.FromAddresses.Contains(sender))
            {
                rule.Conditions.FromAddresses.Add(sender);
            }

            if (rule.Actions.MoveToFolder == null || rule.Actions.MoveToFolder != senderFolder.Id)
            {
                rule.Actions.MoveToFolder = senderFolder.Id;
            }

            rule.Actions.StopProcessingRules = true;

            
            service.UpdateInboxRules(new RuleOperation[] { op }, true);
            
            return rule;
        }

        public static Folder GetSenderFolder(ExchangeService service,Folder parent, string sender)
        {
            Folder folder = null;
            var folderName = string.Join(".", sender.Split('.').Reverse());

            folder = parent.FindFolders(new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, folderName), new FolderView(1)).SingleOrDefault();

            if (folder == null)
            {
                folder = new Folder(service);
                folder.DisplayName = folderName;
                folder.Save(parent.Id);
            }

            return folder;
        }
    }
}
