using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace ewsfetcher
{
    public class Program
    {
        public static IConfiguration Configuration { get; set; }

        public static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json");

            Configuration = builder.Build();

            try
            {
                var ewsClient = new EWSClient(
                    Configuration["mail_url"], 
                    Configuration["mail_domain"], 
                    Configuration["mail_username"], 
                    Configuration["mail_password"]
                );

                if(ewsClient != null)
                {
                    // Get last 3 days emails
                    var timeLimit = new TimeSpan(-3, 0, 0, 0);
                    
                    var inboxItemFetcher = ewsClient.Fetch(timeLimit);
                    inboxItemFetcher.Wait();
                    var inboxItems = inboxItemFetcher.Result;

                    foreach(var inboxItem in inboxItems)
                    {
                        var messageFetcher = ewsClient.LoadMessage(inboxItem.Id);
                        messageFetcher.Wait();
                        var message = messageFetcher.Result;

                        var receivedTime = message.DateTimeReceived;
                        string subject = message.Subject;
                        string senderName = message.From.Name;
                        string senderEmail = message.From.Address;
                        string messageContent = message.Body;
                        bool isRead = message.IsRead;

                        Console.WriteLine("Received email with detail as follow:");
                        Console.WriteLine("  - received: " + receivedTime.ToShortTimeString());
                        Console.WriteLine("  - subject: " + subject);
                        Console.WriteLine("  - sender: " + senderName + " (" + senderEmail + ")");
                        Console.WriteLine("  - content: " + messageContent);
                        Console.WriteLine("  - read: " + isRead.ToString());
                        Console.WriteLine();
                        Console.WriteLine();
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("An error occured(" + ex.Message + "), with stack trace: " + ex.StackTrace);
            }
        }
    }
}
