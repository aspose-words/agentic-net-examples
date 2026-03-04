using System;
using Aspose.Words;

namespace DocumentConversionExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCX file.
            string sourcePath = @"C:\Input\SampleDocument.docx";

            // Path where the resulting PDF will be saved.
            string outputPath = @"C:\Output\SampleDocument.pdf";

            // Load the DOCX document from the file system.
            Document doc = new Document(sourcePath);

            // Save the loaded document as PDF. The format is inferred from the file extension.
            doc.Save(outputPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
