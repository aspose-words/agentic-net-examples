using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocmToPdfConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCM file.
            string inputPath = @"C:\Input\SampleDocument.docm";

            // Path where the resulting PDF will be saved.
            string outputPath = @"C:\Output\SampleDocument.pdf";

            // Load the DOCM file. The Document(string) constructor automatically detects the format.
            Document doc = new Document(inputPath);

            // Save the document as PDF. Using Save(string, SaveFormat) ensures the correct format.
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
