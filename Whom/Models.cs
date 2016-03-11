using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Whom
{
    public class Sender
    {
        public string name { get; internal set; }
        public List<Message> messages { get; internal set; }

        public string subject { get { return messages.First().subject; } }
        public DateTime recieved { get { return messages.First().recieved; } }
        public string body { get { return messages.First().body; } }
        
        public Sender()
        {
            messages = new List<Message>();
        }

    }

    public class Message
    {
        public string body { get; internal set; }
        public DateTime recieved { get; internal set; }
        public string subject { get; internal set; }
        
    }
}
