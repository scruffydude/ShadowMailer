using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShadowMailer
{
    class Report
    {
        private string _name;
        private string _from;
        private string _subject;
        private DateTime _hour;

        private string[] _distroList;
        private string[] _attachments;
        private string[] _embeddedImages;
        private string[] _body;

        public Report(string name, string[] distroList, string[] attachments, string[]embeddedImages, string[] body, string subject = "Default Subject Setup", string from = "Mail@chewy.com")
        {
            _name = name;
            _from = from;
            _subject = subject;
            _body = body;
            _distroList = distroList;
            _attachments = attachments;
            _embeddedImages = embeddedImages;

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
    }
}
