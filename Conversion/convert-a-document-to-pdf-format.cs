using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source document (any supported format, e.g., DOCX).
            string inputFile = @"C:\Docs\SourceDocument.docx";

            // Path where the PDF will be saved. The .pdf extension tells Aspose.Words to save in PDF format.
            string outputFile = @"C:\Docs\ConvertedDocument.pdf";

            // Load the existing document from the file system.
            Document doc = new Document(inputFile);

            // Save the document as PDF. The format is inferred from the file extension.
            doc.Save(outputFile);

            Console.WriteLine("Document successfully converted to PDF.");
        }
    }
}
