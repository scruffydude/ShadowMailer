using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShadowMailer
{
    public class Report
    {
        private string _name;
        private string _from;
        private string _subject;
        private int _hour;
        private string _externalExecutionPath;
        private List<ExcelImageSource> _sources;

        private string[] _distroList;
        private string[] _attachments;
        private string[] _embeddedImages;
        private string[] _body;
        private string _layout;

        public Report(string name, string[] distroList, string[] attachments, string[]embeddedImages, string[] body, string subject = "Default Subject Setup", string from = "Mail@chewy.com", string layout = "", string externalExecutionPath = "UNDEFINED", int hour = -1)
        {
            _name = name;
            _from = from;
            _subject = subject;
            _body = body;
            _layout = layout;
            _distroList = distroList;
            _attachments = attachments;
            _embeddedImages = embeddedImages;
            _externalExecutionPath = externalExecutionPath;
            _hour = hour;

        }
        public string ReportName
        {
            get { return _name; }
            set { _name = value; }
        }
        public string SubjectLine
        {
            get { return _subject; }
            set { _subject = value; }
        }
        public string[] DistroList
        {
            get { return _distroList; }
            set { _distroList = value; }
        }
        public string[] Attachments
        {
            get { return _attachments; }
            set { _attachments = value; }
        }
        public string[] Images
        {
            get { return _embeddedImages; }
            set { _embeddedImages = value; }
        }
        public string[] BodyText
        {
            get { return _body; }
            set { _body = value; }
        }
        public string Sender
        {
            get { return _from; }
            set { _from = value; }
        }
        public string Layout
        {
            get { return _layout; }
            set { _layout = value; }
        }
        public string ExecutablePath
        {
            get { return _externalExecutionPath; }
            set { _externalExecutionPath = value; }
        }
        public int RunHour
        {
            get { return _hour; }
            set { _hour = value; }
        }

    }
}
