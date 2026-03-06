using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOT (Word template) file.
        string inputPath = "InputTemplate.dot";

        // Path where the resulting PDF will be saved.
        string outputPath = "OutputDocument.pdf";

        // Load the DOT document from the file system.
        Document doc = new Document(inputPath);

        // Save the loaded document as PDF.
        // The format is automatically determined from the .pdf extension.
        doc.Save(outputPath);
    }
}
