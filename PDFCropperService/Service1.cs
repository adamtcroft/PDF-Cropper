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
        static int running = 0;
        static int runCount = 0;

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
<<<<<<< HEAD
            FileSystemWatcher watcher = new FileSystemWatcher
            {
                Path = "C:\\Users\\adamt\\Downloads",
                Filter = "label.pdf",
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.FileName | NotifyFilters.LastAccess
            };
=======
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = "C:\\Users\\adamt\\Downloads";
            watcher.Filter = "label.pdf";
            //watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.FileName | NotifyFilters.LastAccess;
>>>>>>> 2d6df4d4df5b416e1e2853298e81813570e7dc40
            watcher.Changed += new FileSystemEventHandler(Watcher_Changed); 
            watcher.Created += new FileSystemEventHandler(Watcher_Changed);
            watcher.EnableRaisingEvents = true;
        }

        private void Watcher_Changed(object sender, FileSystemEventArgs e)
        {
            if (runCount == 0)
            {
                string directory = "C:\\Users\\adamt\\Downloads\\";
                string oldFile = "label.pdf";
                DateTime today = DateTime.Today;
                string date = today.ToString("MMddyyyy");
                string year = today.ToString("yyyy");
                string gDriveDirectory = "C:\\Users\\adamt\\Google Drive\\"+ year + " Taxes\\";
                string eBayName = "eBay-USPS-" + date + ".pdf";
                string newFile = directory + eBayName; 

                cropPage(directory, oldFile, date, year, gDriveDirectory, newFile, eBayName);
                directoryCheck(gDriveDirectory);
                fileCheck(gDriveDirectory, newFile, eBayName);

                runCount++;
            }
            else
            {
                runCount = 0;
            }
        }

        private void cropPage(string directory, string oldFile, string date,
                            string year, string gDriveDirectory, string newFile, string eBayName)
        {
            string file = directory + oldFile;
            oldFile = file;
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

        public void directoryCheck(string gDriveDirectory)
        {
            Console.WriteLine("Made it in directory check");
            if (!Directory.Exists(gDriveDirectory))
            {
                Directory.CreateDirectory(gDriveDirectory);
            }
        }

        public void fileCheck(string gDriveDirectory, string newFile, string eBayName)
        {
            Console.WriteLine("Made it in File Check");

            if (File.Exists(gDriveDirectory + eBayName))
            {
                Console.WriteLine("File exists already");

                int filecount = 1;

                foreach(string filename in Directory.GetFiles(gDriveDirectory))
                {
                    if (filename.Contains(eBayName.Substring(0, 18)))
                    {
                        filecount++;
                        Console.WriteLine(filecount);
                        Console.WriteLine(eBayName.Substring(0, 18));
                    }
                }
                File.Move(newFile, gDriveDirectory + eBayName.Substring(0, 18) + "-pg" + filecount + ".pdf");
            }
            else
            {
                File.Move(newFile, gDriveDirectory + eBayName);
            }
        }

        protected override void OnStop()
        {
        }
    }
}
