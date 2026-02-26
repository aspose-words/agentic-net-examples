using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = "Input.doc";

        // Path to the destination MHTML file.
        string outputPath = "Output.mht";

        // Load the DOC document from the file system.
        Document doc = new Document(inputPath);

        // Save the loaded document in MHTML (Web archive) format.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
