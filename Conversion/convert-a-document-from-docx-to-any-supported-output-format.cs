using System;
using Aspose.Words;

namespace DocumentConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path to the converted file. Change the extension and SaveFormat as needed.
            string outputPath = @"C:\Docs\ConvertedDocument.pdf";

            // Load the DOCX document from the file system.
            Document doc = new Document(inputPath);

            // Save the document in the desired format (PDF in this example).
            // The Save method overload with (string, SaveFormat) follows the provided rule.
            doc.Save(outputPath, SaveFormat.Pdf);

            // Optional: inform the user that conversion succeeded.
            Console.WriteLine($"Document converted successfully to: {outputPath}");
        }
    }
}
