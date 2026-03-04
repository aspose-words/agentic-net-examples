using System;
using Aspose.Words;

class DocToPdfConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"C:\Input\SampleDocument.doc";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Output\SampleDocument.pdf";

        // Load the existing DOC document.
        Document doc = new Document(sourcePath);

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save(outputPath);

        Console.WriteLine("Conversion completed successfully.");
    }
}
