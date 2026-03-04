using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputFile = "input.doc";

        // Path where the MHTML file will be saved.
        string outputFile = "output.mht";

        // Load the DOC document from the file system.
        Document doc = new Document(inputFile);

        // Save the loaded document in MHTML (Web archive) format.
        doc.Save(outputFile, SaveFormat.Mhtml);
    }
}
