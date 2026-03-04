using System;
using Aspose.Words;
using Aspose.Words.Saving;

class LoadDocAndConvert
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourceDocPath = @"C:\Docs\input.doc";

        // Path to the destination file (example: PDF).
        string destinationPath = @"C:\Docs\output.pdf";

        // Load the document. The constructor automatically detects the DOC format.
        Document doc = new Document(sourceDocPath);

        // Save the loaded document in the desired format.
        doc.Save(destinationPath, SaveFormat.Pdf);
    }
}
