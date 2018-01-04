using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Marhsal = System.Runtime.InteropServices;
using System.Security.Principal;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;

namespace ShadowMailer
{
    class Program
    {
        private static Logger log = LogManager.GetCurrentClassLogger();
        private static bool testing = false;
        private static bool debugMode = false;
        private static readonly string reportXmlList = @"C:\Users\camos\source\repos\ShadowMailer\ShadowMailer\ReportsList.xml";
        private static string[] testImages = { @"C:\Users\camos\Desktop\test.jpg", @"C:\Users\camos\Desktop\test2.jpg", @"C:\Users\camos\Desktop\test3.jpg", @"C:\Users\camos\Desktop\test.jpg", @"C:\Users\camos\Desktop\test.jpg", @"C:\Users\camos\Desktop\test2.jpg", @"C:\Users\camos\Desktop\test3.jpg", @"C:\Users\camos\Desktop\test.jpg" };
        private static string[] testBody = { @"C:\Users\camos\Desktop\test.jpg", @"C:\Users\camos\Desktop\test2.jpg", @"C:\Users\camos\Desktop\test3.jpg", @"C:\Users\camos\Desktop\test.jpg" };

        static void Main(string[] args)
        {
            log.Info("Application Shadow Mailer Started by {0}", WindowsIdentity.GetCurrent().Name);

            List<Report> Reports = new List<Report>();

            Reports = getReports();

            if(testing)
            {
                log.Warn("Testing Mode activated may cause instability please review Build");

            }
            else if(!debugMode)
            {
                log.Info("Debug disabled continue checking processing flags");

                foreach(Report currentReport in Reports)
                {
                    SendMail(currentReport.Sender, currentReport.DistroList, currentReport.SubjectLine, currentReport.BodyText[0], testImages);
                }
            }

            try
            {
                File.WriteAllText(reportXmlList, buildXmlFile(Reports));
                log.Info("Reports XML Written back to disk.");
            }
            catch
            {
                log.Fatal("Unable to write XML File to disk.");
            }
            LogManager.Flush();
        }


        public static void SendMail(string from, string[] toList, string subject, string body, string[] ImageLocations)
        {
            var SMTPClient = new SmtpClient("smtp.chewy.local", 587);

            MailMessage message = new MailMessage();
            message.IsBodyHtml = true;
            message.From = new MailAddress(from);

            foreach (string recepiant in toList)
            {
                message.To.Add(recepiant);
            }

            message.Subject = subject;

            message.AlternateViews.Add(getBody(ImageLocations, body));
            SMTPClient.Send(message);
            log.Info("Mail Sent");
        }

        private static AlternateView getBody(string[] filePaths, string body)
        {
            AlternateView alternateView = null;
            string htmlBody = "<html>" + body;
            int i = 0;
            List<LinkedResource> resources = new List<LinkedResource>();
            htmlBody += @"<h1>This is a test</h1></br></br><table ><tr>";
            foreach (string filePath in filePaths)
            {
                if (i%2 == 0) {
                    htmlBody += @"</tr><tr>";
                }
                LinkedResource res = new LinkedResource(filePath);
                res.ContentId = Guid.NewGuid().ToString();
                htmlBody +=@"<td ><img src='cid:" + res.ContentId + @"' /></td>";
                resources.Add(res);
                i++;
                
            }
            htmlBody += "</tr></table>";
            alternateView = AlternateView.CreateAlternateViewFromString(htmlBody, null, MediaTypeNames.Text.Html);
            resources.ForEach(x => alternateView.LinkedResources.Add(x));
            return alternateView;
        }

        private static string buildXmlFile(List<Report> reportsList)
        {
            XmlDocument reports = new XmlDocument();

            XmlNode reportsRoot = reports.CreateElement("Reports");
            reports.AppendChild(reportsRoot);


            foreach (Report report in reportsList)
            {

                XmlNode ReportParent = reports.CreateElement(report.ReportName);
                reportsRoot.AppendChild(ReportParent);
                CreateNewChildXmlNode(reports, ReportParent, "ReportName", report.ReportName.ToString());
                CreateNewChildXmlNode(reports, ReportParent, "SubjectLine", report.SubjectLine.ToString());
                CreateNewChildXmlNode(reports, ReportParent, "Recipiants", report.DistroList[0].ToString());
                foreach (string recepiant in report.DistroList)
                {
                    CreateNewChildXmlNode(reports, ReportParent.LastChild, "Recipiant", recepiant.ToString());
                }

            }

            return reports.InnerXml;
        }

        public static void CreateNewChildXmlNode(XmlDocument document, XmlNode parentNode, string elementName, object value)
        {
            XmlNode node = document.CreateElement(elementName);
            node.AppendChild(document.CreateTextNode(value.ToString()));
            parentNode.AppendChild(node);
        }

        public static List<Report> getReports()
        {
            List<Report> ReportList = new List<Report>();

            string[] DefaultDistro = { "camos@chewy.com", "jlairson@chewy.com" };
            string[] DefaultAttachments = { };

            try
            {
                XmlDocument ReportMaster = new XmlDocument();
                ReportMaster.Load(reportXmlList);
                Report temp = new Report("Default", DefaultDistro,  DefaultAttachments, testImages, testBody);

                foreach (XmlNode Report in ReportMaster.DocumentElement.ChildNodes)
                {
                    temp = new Report("Default", DefaultDistro, DefaultAttachments, testImages, testBody);
                    foreach (XmlNode reportInfo in Report.ChildNodes)
                    {
                        switch (reportInfo.Name)
                        {
                            case "ReportName":
                                temp.ReportName = reportInfo.InnerText;
                                break;
                            case "SubjectLine":
                                temp.SubjectLine = reportInfo.InnerText;
                                break;
                            default:
                                log.Warn("XML Node not found: {0}", reportInfo.Name);
                                break;
                        }
                    }

                    ReportList.Add(temp);
                    log.Info("{0} Warehouse added to list of warehouses", temp.ReportName);
                }
            }
            catch
            {
                Report test = new Report("TestReport", DefaultDistro, DefaultAttachments, testImages, testBody);
                ReportList.Add(test);


            }

            return ReportList;
        }

    }
}
