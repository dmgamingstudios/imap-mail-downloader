using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mail_Downloader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Mail Downloader...");
            Console.WriteLine("Please enter the server ip..");
            Config.Server = Console.ReadLine();
            if(Config.Server.Contains("imap"))
            {
                Console.WriteLine("Detected the imap protocol, setting protocol in config");
                Config.Proc = EAGetMail.ServerProtocol.Imap4;
            }
            else if(Config.Server.Contains("pop"))
            {
                Console.WriteLine("Detected the pop protocol, setting protocol in config");
                Config.Proc = EAGetMail.ServerProtocol.Pop3;
            }
            else
            {
d:              Console.WriteLine("Could not auto-detect protocol, IMAP/POP?");
                string a = Console.ReadLine();
                if (a == "imap")
                    Config.Proc = EAGetMail.ServerProtocol.Imap4;
                else if (a == "pop")
                    Config.Proc = EAGetMail.ServerProtocol.Pop3;
                else
                {
                    Console.WriteLine("Did not understand");
                    Console.Clear();
                    goto d;
                }
            }
            Console.WriteLine("Port?");
            Config.Port = int.Parse(Console.ReadLine());
            Console.WriteLine("Use SSL? Y/N");
            string y = Console.ReadLine();
            if (y == "Y")
                Config.ssl = true;
            else
                Config.ssl = false;
            Console.WriteLine("Username");
            Config.User = Console.ReadLine();
            Console.WriteLine("Password");
            Config.Password = Console.ReadLine();
            Console.Clear();
m:          Console.WriteLine("All set! Start Download? Y/N");
            string r = Console.ReadLine();
            if (r == "Y")
            {
                Download D = new Download();
                if(Config.Proc == EAGetMail.ServerProtocol.Imap4)
                    D.DoDownloadIMAP();
                else
                    D.DoDownloadPOP();
            }
            else if(r == "N")
            {
                Console.Clear();
                goto m;
            }
            Console.ReadLine();
        }
    }
}