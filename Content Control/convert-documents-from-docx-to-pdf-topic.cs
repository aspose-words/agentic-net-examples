using System;
using Aspose.Words;

namespace DocxToPdfConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOCX file path.
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Output PDF file path.
            string outputPath = @"C:\Docs\SampleDocument.pdf";

            // Load the existing DOCX document.
            Document doc = new Document(inputPath);

            // Save the document as PDF. The file extension determines the format.
            doc.Save(outputPath);
        }
    }
}
