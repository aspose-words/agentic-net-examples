using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = "input.docx";

        // Path where the resulting PDF will be saved.
        string outputFile = "output.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputFile);

        // Save the loaded document as PDF.
        // The format is automatically determined from the ".pdf" extension.
        doc.Save(outputFile);
    }
}
