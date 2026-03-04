using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCM file.
            string inputFile = @"C:\Docs\SourceDocument.docm";

            // Path to the target MHTML file. The .mhtml extension tells Aspose.Words to save in MHTML format.
            string outputFile = @"C:\Docs\ConvertedDocument.mhtml";

            // Load the DOCM document.
            Document doc = new Document(inputFile);

            // Save the document as MHTML.
            doc.Save(outputFile, SaveFormat.Mhtml);
        }
    }
}
