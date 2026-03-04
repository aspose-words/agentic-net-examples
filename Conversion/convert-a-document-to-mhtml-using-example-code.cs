using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToMhtml
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words)
        string inputFile = @"C:\Docs\SampleDocument.docx";

        // Path where the MHTML file will be saved
        string outputFile = @"C:\Docs\SampleDocument.mht";

        // Load the document from the file system
        Document doc = new Document(inputFile);

        // Save the document in MHTML format
        // The Save method overload (string, SaveFormat) determines the format from the enum value.
        doc.Save(outputFile, SaveFormat.Mhtml);
    }
}
