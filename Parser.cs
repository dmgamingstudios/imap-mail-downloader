using EAGetMail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Mail_Downloader
{
    class Parser
    {
        internal static void ConvertMailToHtml(string fileName)
        {
            try
            {
                int pos = fileName.LastIndexOf(".");
                string mainName = fileName.Substring(0, pos);
                string htmlName = mainName + ".htm";
                string tempFolder = mainName;
                if (!File.Exists(htmlName))
                {
                    _GenerateHtmlForEmail(htmlName, fileName, tempFolder);
                }
                File.Delete(fileName);
            }
            catch (Exception ep)
            {
                
            }
        }
        private static string _FormatHtmlTag(string src)
        {
            src = src.Replace(">", "&gt;");
            src = src.Replace("<", "&lt;");
            return src;
        }
        private static void _GenerateHtmlForEmail(string htmlName, string emlFile,
                            string tempFolder)
        {
            Mail oMail = new Mail("TryIt");
            oMail.Load(emlFile, false);
            if (oMail.IsEncrypted)
            {
                try
                {
                    oMail = oMail.Decrypt(null);
                }
                catch (Exception ep)
                {
                    oMail.Load(emlFile, false);
                }
            }
            if (oMail.IsSigned)
            {
                try
                {
                    Certificate cert = oMail.VerifySignature();
                }
                catch (Exception ep)
                {
                    
                }
            }
            oMail.DecodeTNEF();
            string html = oMail.HtmlBody;
            StringBuilder hdr = new StringBuilder();
            hdr.Append("<font face=\"Courier New,Arial\" size=2>");
            hdr.Append("<b>From:</b> " + _FormatHtmlTag(oMail.From.ToString()) + "<br>");
            MailAddress[] addrs = oMail.To;
            int count = addrs.Length;
            if (count > 0)
            {
                hdr.Append("<b>To:</b> ");
                for (int i = 0; i < count; i++)
                {
                    hdr.Append(_FormatHtmlTag(addrs[i].ToString()));
                    if (i < count - 1)
                    {
                        hdr.Append(";");
                    }
                }
                hdr.Append("<br>");
            }
            addrs = oMail.Cc;
            count = addrs.Length;
            if (count > 0)
            {
                hdr.Append("<b>Cc:</b> ");
                for (int i = 0; i < count; i++)
                {
                    hdr.Append(_FormatHtmlTag(addrs[i].ToString()));
                    if (i < count - 1)
                    {
                        hdr.Append(";");
                    }
                }
                hdr.Append("<br>");
            }
            hdr.Append(String.Format("<b>Subject:</b>{0}<br>\r\n",
                _FormatHtmlTag(oMail.Subject)));
            Attachment[] atts = oMail.Attachments;
            count = atts.Length;
            if (count > 0)
            {
                if (!Directory.Exists(tempFolder))
                    Directory.CreateDirectory(tempFolder);

                hdr.Append("<b>Attachments:</b>");
                for (int i = 0; i < count; i++)
                {
                    Attachment att = atts[i];
                    string attname = String.Format("{0}\\{1}", tempFolder, att.Name);
                    att.SaveAs(attname, true);
                    hdr.Append(String.Format("<a href=\"{0}\" target=\"_blank\">{1}</a> ",
                            attname, att.Name));
                    if (att.ContentID.Length > 0)
                    {
                        html = html.Replace("cid:" + att.ContentID, attname);
                    }
                    else if (String.Compare(att.ContentType, 0, "image/", 0,
                                "image/".Length, true) == 0)
                    {
                        html = html + String.Format("<hr><img src=\"{0}\">", attname);
                    }
                }
            }
            Regex reg = new Regex("(<meta[^>]*charset[ \t]*=[ \t\"]*)([^<> \r\n\"]*)",
                RegexOptions.Multiline | RegexOptions.IgnoreCase);
            html = reg.Replace(html, "$1utf-8");
            if (!reg.IsMatch(html))
            {
                hdr.Insert(0,
                    "<meta HTTP-EQUIV=\"Content-Type\" Content=\"text-html; charset=utf-8\">");
            }
            html = hdr.ToString() + "<hr>" + html;
            FileStream fs = new FileStream(htmlName, FileMode.Create,
                FileAccess.Write, FileShare.None);
            byte[] data = System.Text.UTF8Encoding.UTF8.GetBytes(html);
            fs.Write(data, 0, data.Length);
            fs.Close();
            oMail.Clear();
        }
    }
}