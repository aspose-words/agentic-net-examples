using System;
using Aspose.Words;

namespace AsposeWordsPdfToDocExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF file.
            string pdfFilePath = @"C:\Input\sample.pdf";

            // Path where the converted DOC file will be saved.
            string docFilePath = @"C:\Output\sample.doc";

            // Load the PDF document. The Document constructor automatically detects the format.
            Document pdfDocument = new Document(pdfFilePath);

            // Save the document as a legacy Word .doc file.
            // The file extension determines the save format, but we can also specify it explicitly.
            pdfDocument.Save(docFilePath, SaveFormat.Doc);
        }
    }
}
