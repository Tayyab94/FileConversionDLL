using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryConvert
{
    public class Class1
    {
        public static void add()
        {
            Console.WriteLine("This is Add function");
        }


       public static Microsoft.Office.Interop.Word.Document wordDoc { get; set; }

        public static void Convertor()
        {
          //Microsoft.Office.Interop.Word.Document wordDoc { get; set; }
             Console.WriteLine("Please wait.....");
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            wordDoc = app.Documents.Open(@"C:\Users\Hp\Desktop\TaskTest\TextFile.docx");
            wordDoc.ExportAsFixedFormat(@"C:\Users\Hp\Desktop\TaskTest\FileConversion\WordConv\TextFile.pdf", Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);

            app.Quit();


            //NameSpace is using System.Runtime.Interopservices;
            Marshal.ReleaseComObject(wordDoc);
            Marshal.ReleaseComObject(app);

            Console.WriteLine("Thank you for Word Document Conversion");
        }

        public static void pdfToWord()
        {
            Console.WriteLine("Please wait.....");
            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            f.OpenPdf(@"C:\Users\Hp\Desktop\TaskTest\TextFile.pdf");
            f.ToWord(@"C:\Users\Hp\Desktop\TaskTest\FileConversion\WordConv\TextFile.doc");

            Console.WriteLine("Thank you for Pdf Document Conversion");
        }


        public static void Convertor1(string sourcePath,string fname)
        {
            //Microsoft.Office.Interop.Word.Document wordDoc { get; set; }
            Console.WriteLine("Please wait.....");
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            // wordDoc = app.Documents.Open(@"C:\Users\Hp\Desktop\TaskTest\TextFile.docx");

            wordDoc = app.Documents.Open($"{sourcePath}{fname}");
            wordDoc.ExportAsFixedFormat(@"C:\Users\Hp\Desktop\TaskTest\FileConversion\WordConv\TextFile.pdf", Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);

            app.Quit();

            //NameSpace is using System.Runtime.Interopservices;
            Marshal.ReleaseComObject(wordDoc);
            Marshal.ReleaseComObject(app);

            Console.WriteLine("Thank you for Word Document Conversion");
        }

        public static void pdfToWord1(string sourcePath, string fname)
        {
            Console.WriteLine("Please wait.....");
            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            // f.OpenPdf(@"C:\Users\Hp\Desktop\TaskTest\TextFile.pdf");
            f.OpenPdf($"{sourcePath}{fname}");
            f.ToWord(@"C:\Users\Hp\Desktop\TaskTest\FileConversion\WordConv\TextFile.doc");

            Console.WriteLine("Thank you for Pdf Document Conversion");
        }
    }
}
