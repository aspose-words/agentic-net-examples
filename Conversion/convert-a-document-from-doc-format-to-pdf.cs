using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputFile = "input.doc";

        // Path where the resulting PDF will be saved.
        string outputFile = "output.pdf";

        // Load the DOC document from the file system.
        Document doc = new Document(inputFile);

        // Save the loaded document as PDF.
        // The format is automatically determined from the ".pdf" extension.
        doc.Save(outputFile);
    }
}
