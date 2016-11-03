using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EAGetMail;
namespace Mail_Downloader
{
    class Config
    {
        internal static string Server = "";
        internal static string User = "";
        internal static string Password = "";
        internal static ServerProtocol Proc;
        internal static bool ssl = false;
        internal static int Port;
    }
}