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
using System.Diagnostics;

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
        private static string testLayout = "BIIBIIIIF";
        private static string testcommand = @"\\cfc1afs01\Operations-Analytics\Projects\CSharp\labor-plan\Solutions\mcoapp\mcoapp\bin\Release\mcoapp.exe";
        private static string testworkbook = @"C:\Users\camos\Desktop\testExcel.xlsx";
        private static string testRange = "B5";

        static void Main(string[] args)
        {
            log.Info("Application Shadow Mailer Started by {0}", WindowsIdentity.GetCurrent().Name);

            List<Report> Reports = new List<Report>();

            Reports = getReports();

            if(testing)
            {
                CollectImageFromExcelRange(testworkbook, testRange);
                log.Warn("Testing Mode activated may cause instability please review Build");

            }
            else if(!debugMode)
            {
                log.Info("Debug disabled continue checking processing flags");

                foreach(Report currentReport in Reports)
                {
                    if (currentReport.RunHour == System.DateTime.Now.Hour)
                    {
                        if (currentReport.ExecutablePath != "UNDEFINED")
                        {
                            RunExternalDataManipulation(currentReport.ExecutablePath);
                        }
                        SendMail(currentReport.Sender, currentReport.DistroList, currentReport.SubjectLine, currentReport.BodyText, testImages, currentReport.Layout);
                    }
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

        public static void CollectImageFromExcelRange(string workbookPath, string picRange)
        {

        }

        public static void SendMail(string from, string[] toList, string subject, string[] body, string[] ImageLocations, string layoutControl)
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

            message.AlternateViews.Add(buildBodyTable(ImageLocations, body, layoutControl));//getBody(ImageLocations, body, layoutControl));
            SMTPClient.Send(message);
            log.Info("Mail Sent");
        }

        public static AlternateView buildBodyTable(string[] filepaths, string[] body, string layoutControl)
        {
            AlternateView alternateView = null;
            List<LinkedResource> resources = new List<LinkedResource>();


            string htmlBody = "<html><body>";
            htmlBody += @"<h1>This is a test</h1></br></br><table >";

            //these are counters are used to cycle through the element refernces in the body and images setup
            int b = 0;
            int i = 0;

            foreach(char c in layoutControl)
            {
                switch (c)
                {
                    case 'B':
                        htmlBody += "<tr><td colspan=\"2\" >" +body[b]+ "</td></tr>";
                        b++;
                        break;
                    case 'I':
                        if (i % 2 == 0)
                            htmlBody += @"<tr>";

                        resources.Add(getImage(filepaths[i], ref htmlBody));

                        if (i % 2 != 0)
                            htmlBody += @"</tr>";

                        i++;
                        break;
                    case 'F':
                        resources.Add(getImage(filepaths[i], ref htmlBody));
                        i++;
                        break;
                    default:
                        log.Warn("Character Designation not found skipping....");
                        break;
                }
            }

            htmlBody += "</table></html>";
            alternateView = AlternateView.CreateAlternateViewFromString(htmlBody, null, MediaTypeNames.Text.Html);
            resources.ForEach(x => alternateView.LinkedResources.Add(x));

            return alternateView;
        }

        private static LinkedResource getImage(string filePath, ref string htmlBody)
        {
            LinkedResource res = new LinkedResource(filePath);
            res.ContentId = Guid.NewGuid().ToString();
            htmlBody += @"<td ><img src='cid:" + res.ContentId + @"' /></td>";
            return res;

        }
        private static string buildXmlFile(List<Report> reportsList)
        {
            XmlDocument reports = new XmlDocument();

            XmlNode reportsRoot = reports.CreateElement("Reports");
            reports.AppendChild(reportsRoot);


            foreach (Report report in reportsList)
            {

                XmlNode CurrentReportParent = reports.CreateElement(report.ReportName);
                reportsRoot.AppendChild(CurrentReportParent);
                CreateNewChildXmlNode(reports, CurrentReportParent, "ReportName", report.ReportName.ToString());
                CreateNewChildXmlNode(reports, CurrentReportParent, "SubjectLine", report.SubjectLine.ToString());
                CreateNewChildXmlNode(reports, CurrentReportParent, "Sender", report.Sender.ToString());

                CreateChildNodeFromArray("Recipient", report.DistroList, reports, CurrentReportParent);

                CreateChildNodeFromArray("BodyText", report.BodyText, reports, CurrentReportParent);

                CreateChildNodeFromArray("Image", report.Images, reports, CurrentReportParent);

                CreateNewChildXmlNode(reports, CurrentReportParent, "Layout", report.Layout.ToString());
                CreateNewChildXmlNode(reports, CurrentReportParent, "ExternalExecutable", report.ExecutablePath.ToString());
                CreateNewChildXmlNode(reports, CurrentReportParent, "RunHour", report.RunHour.ToString());

            }

            return reports.InnerXml;
        }

        public static void CreateChildNodeFromArray(string childNodeName, string[] elementArray, XmlDocument document, XmlNode parentNode)
        {
            XmlNode newNode = document.CreateElement(childNodeName+"s");
            parentNode.AppendChild(newNode);

            foreach(string element in elementArray)
            {
                CreateNewChildXmlNode(document, newNode, childNodeName, element);
            }
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
                    temp = new Report("Default", DefaultDistro, DefaultAttachments, testImages, testBody, layout: testLayout);
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
                            case "Layout":
                                temp.Layout = reportInfo.InnerText;
                                break;
                            case "Recipients":
                                temp.DistroList = populateArrayfromXML(reportInfo);
                                break;
                            case "Sender":
                                temp.Sender = reportInfo.InnerText;
                                break;
                            case "ExternalExecutable":
                                temp.ExecutablePath = reportInfo.InnerText;
                                break;
                            case "RunHour":
                                temp.RunHour = int.Parse(reportInfo.InnerText);
                                break;
                            case "BodyTexts":
                                temp.BodyText = populateArrayfromXML(reportInfo);
                                break;
                            case "Images":
                                temp.Images = populateArrayfromXML(reportInfo);
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
                Report test = new Report("TestReport", DefaultDistro, DefaultAttachments, testImages, testBody, layout: testLayout);
                ReportList.Add(test);
            }

            return ReportList;
        }

        public static string[] populateArrayfromXML(XmlNode parent)
        {
            List<string> templist = new List<string>();

            foreach(XmlNode child  in parent)
            {
                templist.Add(child.InnerText);
            }

            return templist.ToArray();

        }

        public static void RunExternalDataManipulation(string commandInformantion)
        {
            var process  = Process.Start(commandInformantion);
            process.WaitForExit();
        }

    }
}
