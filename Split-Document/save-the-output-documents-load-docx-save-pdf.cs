using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Path where the PDF will be saved.
            string outputPath = @"C:\Docs\SampleDocument.pdf";

            // Load the existing DOCX document.
            Document doc = new Document(inputPath);

            // Save the loaded document as PDF. The format is inferred from the file extension.
            doc.Save(outputPath);
        }
    }
}
