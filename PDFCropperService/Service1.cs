using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace PDFCropperService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        public void OnDebug()
        {
            OnStart(null);
        }

        protected override void OnStart(string[] args)
        {
            string directory = "C:\\Users\\adamt\\Downloads\\";
            string oldFile = "label.pdf";
            DateTime today = DateTime.Today;
            string date = today.ToString("MMddyyyy");

            FileSystemWatcher watcher = new FileSystemWatcher(directory, filter: "label.pdf");

            cropPage(directory, oldFile, date);
        }

        private void cropPage(string directory, string oldFile, string date)
        {
            string file = directory + "label.pdf";
            oldFile = file;
            string newFile = directory + "eBay-USPS-" + date + ".pdf";
            PdfReader reader = new PdfReader(file);

            iTextSharp.text.Rectangle size = new iTextSharp.text.Rectangle(0, 0, 612, 422);

            Document doc = new Document(size);

            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(
                file.Replace(oldFile, newFile), FileMode.Create, FileAccess.Write));

            doc.Open();

            PdfContentByte cb = writer.DirectContent;

            doc.NewPage();

            PdfImportedPage page = writer.GetImportedPage(reader, 1);
            cb.AddTemplate(page, 0, 0);
            doc.Close();
        }

        protected override void OnStop()
        {
        }
    }
}
