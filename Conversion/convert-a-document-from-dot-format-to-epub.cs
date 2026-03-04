using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DotToEpubConverter
{
    static void Main()
    {
        // Path to the source DOT (Word template) file.
        string inputPath = @"C:\Docs\Template.dot";

        // Path where the resulting EPUB file will be saved.
        string outputPath = @"C:\Docs\Converted.epub";

        // Load the DOT document. The Document constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Save the document in EPUB format.
        doc.Save(outputPath, SaveFormat.Epub);
    }
}
