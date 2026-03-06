using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDotToEpub
{
    static void Main()
    {
        // Path to the source DOT (Word template) file.
        string inputPath = "InputTemplate.dot";

        // Path where the resulting EPUB file will be saved.
        string outputPath = "OutputDocument.epub";

        // Load the DOT document.
        Document doc = new Document(inputPath);

        // Save the loaded document in EPUB format.
        doc.Save(outputPath, SaveFormat.Epub);
    }
}
