using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOC file path
        string inputPath = "Input.doc";

        // Desired output MHTML file path
        string outputPath = "Output.mhtml";

        // Load the DOC document from the file system
        Document doc = new Document(inputPath);

        // Save the document in MHTML format using the SaveFormat enumeration
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
