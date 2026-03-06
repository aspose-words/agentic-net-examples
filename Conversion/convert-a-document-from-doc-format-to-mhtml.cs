using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\Sample.doc";

        // Path where the MHTML file will be saved.
        string outputPath = @"C:\Docs\Sample.mhtml";

        // Load the DOC document. The constructor detects the format automatically.
        Document doc = new Document(inputPath);

        // Save the document in MHTML (Web archive) format.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
