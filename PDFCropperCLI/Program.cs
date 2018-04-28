using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PDFCropperCLI
{
    class Program
    {
        static void Main(string[] args)
        {
            string directory = "C:\\Users\\" + Environment.UserName + "\\Downloads\\";
            string oldFile = "label.pdf";
            DateTime today = DateTime.Today;
            string date = today.ToString("MMddyyyy");
            string year = today.ToString("yyyy");
            string gDriveDirectory = "Z:\\Taxes\\" + year + " Taxes\\";
            string eBayName = "eBay-USPS-" + date + ".pdf";
            string newFile = directory + eBayName;

            cropPage(directory, oldFile, date, year, gDriveDirectory, newFile, eBayName);
            directoryCheck(gDriveDirectory);
            fileCheck(gDriveDirectory, newFile, eBayName);
        }

        private static void cropPage(string directory, string oldFile, string date,
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

        private static void directoryCheck(string gDriveDirectory)
        {
            if (!Directory.Exists(gDriveDirectory))
            {
                Directory.CreateDirectory(gDriveDirectory);
            }
        }

        private static void fileCheck(string gDriveDirectory, string newFile, string eBayName)
        {
            try
            {
                File.Move(newFile, gDriveDirectory + eBayName);
                Console.WriteLine("File Created: " + gDriveDirectory + eBayName);
            }
            catch (IOException e)
            {
                if (File.Exists(gDriveDirectory + eBayName))
                {
                    int filecount = 1;

                    foreach (string filename in Directory.GetFiles(gDriveDirectory))
                    {
                        if (filename.Contains(eBayName.Substring(0, 18)))
                        {
                            filecount++;
                        }
                    }
                    File.Move(newFile, gDriveDirectory + eBayName.Substring(0, 18) + "-pg" + filecount + ".pdf");
                    Console.WriteLine("File Created: " + gDriveDirectory + eBayName.Substring(0, 18) + "-pg" + filecount + ".pdf");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
