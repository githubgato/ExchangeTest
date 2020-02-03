using System;
using Microsoft.Exchange.WebServices.Data;


namespace GetMailFromExchange
{
    class Program
    {
        ExchangeService serviceInstance;
        public string ExceptionMessage { get; }
        static void Main(string[] args)
        {
            Console.WriteLine("start!");
            ConnectToExchangeServer();
            //ExchangeRepository();
        }

        static void ConnectToExchangeServer()
        {

            Console.WriteLine("connect to exchangeServer!");

            try
            {
                //ExchangeService exchange = new ExchangeService(ExchangeVersion.);
                //exchange.Credentials = new WebCredentials("svcMailRead@qsc.com", "Qsc123" );
                //exchange.AutodiscoverUrl("svcMailRead@qsc.com");

                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

                //service.UseDefaultCredentials = true;

                //service.Credentials = new WebCredentials("123@abc.com", "abc", "123");
                //service.Credentials = new WebCredentials("123@abc.com", "abc", "123");
                service.Credentials = new WebCredentials("123@abc.com", "abc" );
                service.TraceEnabled = true;
                service.TraceFlags = TraceFlags.All;
                service.AutodiscoverUrl("123@abc.com");
                EmailMessage email = new EmailMessage(service);
                email.ToRecipients.Add("123@abc.com");
                email.Subject = "HelloWorld";
                email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
                email.Send();

                Console.WriteLine("Connected to Exchange Server : " + service.Url.Host);


            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Connecting to Exchange Server!!" + ex.Message);


            }


        }


        static void ExchangeRepository()
        {
            ExchangeService  serviceInstance = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            //Provide the account user name in format 123@abc.com
            serviceInstance.Credentials = new WebCredentials("123@abc.com", "abc", "123");

            try
            {
                // Use Autodiscover to set the URL endpoint.
                // and using a AutodiscoverRedirectionUrlValidationCallback in case of https enabled clod account
                serviceInstance.AutodiscoverUrl("123@abc.com", SslRedirectionCallback);
                Console.WriteLine("Connecting : ");
            }
            catch (Exception ex)
            {
                serviceInstance = null;
                 
                Console.WriteLine("Connected to Exchange Server : " + ex.Message);

            }

        }

        static bool SslRedirectionCallback(string serviceUrl)
        {
            // Return true if the URL is an HTTPS URL.
            return serviceUrl.ToLower().StartsWith("https://");
        }

    }
}
