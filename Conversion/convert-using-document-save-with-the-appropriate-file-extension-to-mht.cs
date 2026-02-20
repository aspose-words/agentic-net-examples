using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocmToMhtml
{
    static void Main()
    {
        // Path to the source DOCM file.
        const string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path to the destination MHTML file. The .mhtml extension tells Aspose.Words the desired format.
        const string outputPath = @"C:\Docs\ConvertedDocument.mhtml";

        // Load the macro‑enabled document.
        Document doc = new Document(inputPath);

        // Save the document as MHTML. The SaveFormat enumeration explicitly specifies the format.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
