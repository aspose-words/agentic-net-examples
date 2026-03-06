using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source document (DOCX, DOC, etc.).
            string inputPath = @"C:\Docs\ExampleDocument.docx";

            // Path where the PDF will be saved.
            string outputPath = @"C:\Docs\ExampleDocument.pdf";

            // Load the existing document from the file system.
            Document doc = new Document(inputPath);

            // Save the document as PDF. The format is inferred from the .pdf extension.
            doc.Save(outputPath);

            Console.WriteLine("Document successfully converted to PDF.");
        }
    }
}
