using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOCX file path
            string inputFile = @"C:\Docs\SampleDocument.docx";

            // Output PDF file path
            string outputFile = @"C:\Docs\SampleDocument.pdf";

            // Load the existing DOCX document (uses Document(string) constructor)
            Document doc = new Document(inputFile);

            // Save the document as PDF (uses Document.Save(string) overload)
            doc.Save(outputFile);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
