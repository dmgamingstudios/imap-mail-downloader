using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EAGetMail;
using System.IO;
using System.Text.RegularExpressions;

namespace Mail_Downloader
{
    class Download
    {
        MailServer Server = new MailServer(Config.Server, Config.User, Config.Password, Config.Proc);
        MailClient Client = new MailClient("TryIt");
        string saveDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Downloaded Mails";
        internal void DoDownloadIMAP()
        {
            Console.WriteLine("Initiating Download...");
            if (!Directory.Exists(saveDir))
                Directory.CreateDirectory(saveDir);
            Server.SSLConnection = Config.ssl;
            Server.Port = Config.Port;
            try
            {
                Console.WriteLine("Connecting...");
                Client.Connect(Server);
                Console.WriteLine("Connected");
                Imap4Folder[] Folders = Client.Imap4Folders;
                for(int da = 0; da < Folders.Length; da++)
                {
                    Imap4Folder Folder = Folders[da];
                    Client.SelectFolder(Folder);
                    MailInfo[] infos = Client.GetMailInfos();
                    for (int i = 0; i < infos.Length; i++)
                    {
                        if (!Directory.Exists(String.Format("{0}\\{1}", saveDir, Folder.Name.ToString())))
                            Directory.CreateDirectory(String.Format("{0}\\{1}", saveDir, Folder.Name.ToString()));
                        MailInfo info = infos[i];
                        Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}; Folder: {3}",
                            info.Index, info.Size, info.UIDL, Folder.Name.ToString());
                        Mail oMail = Client.GetMail(info);
                        try
                        { 
                            Console.WriteLine("From: {0}", oMail.From.ToString());
                            Console.WriteLine("Subject: {0}\r\n", oMail.Subject);
                            string fileName = String.Format("{0}{1}.eml",
                                oMail.Subject.ToString(), i);
                            string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                            Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                            fileName = r.Replace(fileName, "");
                            oMail.SaveAs(saveDir + "\\" + Folder.Name.ToString() + "\\" + fileName, true);
                        }
                        catch(PathTooLongException ex)
                        {
                            string fileName = String.Format("{0}{1}{2}.eml",
                                oMail.From.ToString(), oMail.ReceivedDate.ToString(), i);
                            string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                            Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                            fileName = r.Replace(fileName, "");
                            oMail.SaveAs(saveDir + "\\" + Folder.Name.ToString() + "\\" + fileName, true);
                        }
                    }
                }
            }
            catch(MailServerException ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Client.Quit();
            }
            Console.Clear();
            Console.WriteLine("All Done!! Parsing mails (so they are readable)");
            string[] dirs = Directory.GetDirectories(saveDir);
            for(int a = 0; a < dirs.Length; a++)
            {
                string[] files = Directory.GetFiles(dirs[a], "*.eml");
                for (int i = 0; i < files.Length; i++)
                {
                    Console.WriteLine("Parsing " + files[i]);
                    Parser.ConvertMailToHtml(files[i]);
                }
            }
            Console.Clear();
            Console.WriteLine("All Finished, thanks for using me :D");
            Console.ReadLine();
        }
        internal void DoDownloadPOP()
        {
            Console.WriteLine("Initiating Download...");
            if (!Directory.Exists(saveDir))
                Directory.CreateDirectory(saveDir);
            Server.SSLConnection = Config.ssl;
            Server.Port = Config.Port;
            try
            {
                Console.WriteLine("Connecting...");
                Client.Connect(Server);
                Console.WriteLine("Connected");
                MailInfo[] infos = Client.GetMailInfos();
                for (int i = 0; i < infos.Length; i++)
                {
                    MailInfo info = infos[i];
                    Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL);
                    Mail oMail = Client.GetMail(info);
                    Console.WriteLine("From: {0}", oMail.From.ToString());
                    Console.WriteLine("Subject: {0}\r\n", oMail.Subject);
                    System.DateTime d = System.DateTime.Now;
                    System.Globalization.CultureInfo cur = new
                        System.Globalization.CultureInfo("en-US");
                    string sdate = d.ToString("yyyyMMddHHmmss", cur);
                    string fileName = String.Format("{0}\\{1}{2}{3}{4}.eml",
                        saveDir, oMail.Subject.ToString(), sdate, d.Millisecond.ToString("d3"), i);
                    oMail.SaveAs(fileName, true);
                }
            }
            catch (MailServerException ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Client.Quit();
            }
            Console.Clear();
            Console.WriteLine("All Done!! Parsing mails (so they are readable)");
            string[] dirs = Directory.GetDirectories(saveDir);
            for (int a = 0; a < dirs.Length; a++)
            {
                string[] files = Directory.GetFiles(dirs[a], "*.eml");
                for (int i = 0; i < files.Length; i++)
                {
                    Console.WriteLine("Parsing " + files[i]);
                    Parser.ConvertMailToHtml(files[i]);
                }
            }
            Console.Clear();
            Console.WriteLine("All Finished, thanks for using me :D");
            Console.ReadLine();
        }
    }
}