using System;
using Aspose.Words;

namespace DocmToMhtmlConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCM file.
            string inputPath = @"C:\Docs\SourceDocument.docm";

            // Path where the resulting MHTML file will be saved.
            string outputPath = @"C:\Docs\ConvertedDocument.mht";

            // Load the DOCM document.
            Document doc = new Document(inputPath);

            // Save the document in MHTML (Web archive) format.
            doc.Save(outputPath, SaveFormat.Mhtml);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
