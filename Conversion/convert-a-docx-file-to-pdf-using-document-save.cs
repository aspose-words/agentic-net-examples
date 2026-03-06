using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCX file.
            string inputFile = @"C:\Docs\SampleDocument.docx";

            // Path where the resulting PDF will be saved.
            string outputFile = @"C:\Docs\SampleDocument.pdf";

            // Load the DOCX document from the file system.
            Document doc = new Document(inputFile);

            // Save the loaded document as PDF. The format is inferred from the file extension.
            doc.Save(outputFile);
        }
    }
}
