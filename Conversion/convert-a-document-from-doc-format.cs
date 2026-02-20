using System;
using Aspose.Words;

class DocConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\source.doc";

        // Path where the converted file will be saved (e.g., PDF).
        string outputPath = @"C:\Docs\converted.pdf";

        // Load the DOC document. The constructor automatically detects the format.
        Document document = new Document(inputPath);

        // Save the document in the desired format (PDF in this example).
        document.Save(outputPath, SaveFormat.Pdf);
    }
}
