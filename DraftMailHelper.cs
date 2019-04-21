using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace VstoHelperTest.Helper
{
    class DraftMailHelper
    {
        private string[] _to;
        private string[] _cc;
        private string[] _bcc;
        private string[] _attachmentPath;
        private string _subject;
        private Outlook.Application _outlookApp;
        private Outlook.NameSpace _nameSpace;
        private Outlook.MAPIFolder _drafMailFolder;
        private Outlook.MailItem _targetMail;
        private Outlook.Attachments _attachments;
        private Excel.Worksheet _mailBodyWorksheetTemplate;

        public dynamic SetTo
        {
            set => _to = value is string ? ConverStringToStringArray(value) : value;
        }

        public dynamic SetCc
        {
            set => _cc = value is string ? ConverStringToStringArray(value) : value;
        }

        public dynamic SetBcc
        {
            set => _bcc = value is string ? ConverStringToStringArray(value) : value;
        }

        public dynamic SetAttachmentPath
        {
            set => _attachmentPath = value is string ? ConverStringToStringArray(value) : value;
        }

        public string SetSubject
        {
            set => _subject = value;
        }

        public Excel.Worksheet SetMailBodyWorksheetTemplate
        {
            set => _mailBodyWorksheetTemplate = value;
        }

        public DraftMailHelper()
        {
            try
            {
                _outlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            catch (System.Exception)
            {
                CommonFunctionHelper.ErrorHandling("Please open the Outlook before drafting the mail.");
            }

            _nameSpace = _outlookApp?.GetNamespace("MAPI");
            _drafMailFolder = _nameSpace?.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
            _subject = "";
            _to = new string[] { };
            _cc = new string[] { };
            _bcc = new string[] { };
            _attachmentPath = new string[] { };
        }

        //--------------------------------------------
        // Main Function
        //--------------------------------------------
        public void CreateDraftMailItem()
        {
            _targetMail = (_drafMailFolder.Items.Add(Outlook.OlItemType.olMailItem)).Move(_drafMailFolder) as Outlook.MailItem;
        }

        public void SetRequiredSetting()
        {
            _targetMail.Display();
            _targetMail.Subject = _subject;
            _targetMail.CC = TransStringArrayToString(_cc);
            _targetMail.To = TransStringArrayToString(_to);
            _targetMail.BCC = TransStringArrayToString(_bcc);
            _targetMail.Close(Outlook.OlInspectorClose.olSave);
        }

        public void SetBodyText()
        {
            var usedRangeSource = _mailBodyWorksheetTemplate.UsedRange;
            _targetMail.HTMLBody = RangeToHtml(usedRangeSource);
        }

        public void AddAttachment()
        {
            _attachments = _targetMail.Attachments;
            foreach (var item in _attachmentPath)
                _attachments.Add(item.Trim());
        }

        public void CloseDraftMailItem()
        {
            if (_targetMail != null)
                Marshal.ReleaseComObject(_targetMail);
            if (_attachments != null)
                Marshal.ReleaseComObject(_attachments);
            if (_mailBodyWorksheetTemplate != null)
                Marshal.ReleaseComObject(_mailBodyWorksheetTemplate);
        }

        public void DisposeOutlookApp()
        {
            if (_drafMailFolder != null)
                Marshal.ReleaseComObject(_drafMailFolder);
            if (_nameSpace != null)
                Marshal.ReleaseComObject(_nameSpace);
            if (_outlookApp != null)
                Marshal.ReleaseComObject(_outlookApp);
        }

        //--------------------------------------------
        // Internal Function
        //--------------------------------------------
        private string[] ConverStringToStringArray(string targetString)
        {
            return targetString.Split(',');
        }

        private string TransStringArrayToString(string[] targetArray)
        {
            return targetArray.Aggregate("", (current, item) => current + item + ";");
        }

        private string RangeToHtml(Excel.Range targetRange)
        {
            Excel.Worksheet targetWorksheet = targetRange.Parent;
            var pathTemp = Path.GetTempPath() + "tempHtmlFile.htm";

            targetWorksheet.Copy();

            var excelApp = Globals.ThisWorkbook.Application;
            var tempWorkbook = excelApp.ActiveWorkbook;
            
            tempWorkbook.PublishObjects.Add(Excel.XlSourceType.xlSourceRange,pathTemp, 
                targetWorksheet.Name,targetRange.Address, Excel.XlHtmlType.xlHtmlStatic).Publish();

            tempWorkbook.Saved = true;
            tempWorkbook.Close();

            var text = File.ReadAllText(pathTemp, Encoding.Default);
            var htmlText = text.Replace("align=center x:publishsource=", "align=left x:publishsource=");

            File.Delete(pathTemp);
            return htmlText;
        }
    }
}
