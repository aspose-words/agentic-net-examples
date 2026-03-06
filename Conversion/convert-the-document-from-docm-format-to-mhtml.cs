using System;
using Aspose.Words;

namespace DocmToMhtmlConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCM file.
            string inputPath = @"C:\Docs\SourceDocument.docm";

            // Path where the resulting MHTML file will be saved.
            string outputPath = @"C:\Docs\ConvertedDocument.mht";

            // Load the DOCM document. The Document constructor automatically detects the format.
            Document doc = new Document(inputPath);

            // Save the document in MHTML (Web archive) format.
            // This uses the Document.Save(string, SaveFormat) overload.
            doc.Save(outputPath, SaveFormat.Mhtml);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
