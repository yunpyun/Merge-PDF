using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Document;
using GemBox.Pdf;

namespace MergePdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // лицензирование GemBox
            GemBox.Document.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // исходный массив с файлами
            var fileNames = new string[]
            {
                @"C:\Users\Выймова Елена\Documents\InputDocMerge.docx",
                @"C:\Users\Выймова Елена\Documents\InputPdfMerge.pdf",
                @"C:\Users\Выймова Елена\Documents\InputPdfMerge2.pdf"
            };

            // конвертация из исходного массива doc файлов в pdf
            for (int i = 0; i < fileNames.Length; i++)
            {
                if (fileNames[i].EndsWith(".doc") || fileNames[i].EndsWith(".docx"))
                {
                    fileNames[i] = SaveDoc(fileNames[i]);
                }
            }

            // объединение pdf файлов
            MergeFiles(fileNames);

            Console.WriteLine("Saving pdf done");
            Console.ReadLine();
        }

        static string SaveDoc(string destFileName)
        {
            DocumentModel documentDoc = DocumentModel.Load(destFileName);
            string newName = @"C:\Users\Выймова Елена\Documents\OutputDocMerge.pdf";
            documentDoc.Save(newName);
            return newName;
        }

        static void MergeFiles(string[] fileNames)
        {
            using (PdfDocument documentPdf = new PdfDocument())
            {
                foreach (var fileName in fileNames)
                    using (var source = PdfDocument.Load(fileName))
                        documentPdf.Pages.Kids.AddClone(source.Pages);

                documentPdf.Save(@"C:\Users\Выймова Елена\Documents\MergeFiles.pdf");
            }
        }
    }
}
