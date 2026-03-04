using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentConversionExample
{
    class Program
    {
        static void Main()
        {
            // Define the input DOC file path.
            string inputPath = @"C:\MyDir\Document.doc";

            // Define the output MHTML file path.
            string outputPath = @"C:\ArtifactsDir\Document.mht";

            // Load the DOC document from the file system.
            Document doc = new Document(inputPath);

            // Save the document in MHTML (Web archive) format.
            // The Save method with a file name and SaveFormat enumerator is used as per the provided rules.
            doc.Save(outputPath, SaveFormat.Mhtml);

            // Optional: confirm that the file was created.
            Console.WriteLine($"Document converted to MHTML and saved at: {outputPath}");
        }
    }
}
