using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DotToMhtmlConverter
{
    static void Main()
    {
        // Path to the source DOT (Word template) file.
        string inputPath = @"C:\Docs\Template.dot";

        // Path where the resulting MHTML file will be saved.
        string outputPath = @"C:\Docs\Template.mht";

        // Load the DOT document. The constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Save the document in MHTML format.
        // Using the overload that accepts a file name and a SaveFormat enumeration.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
