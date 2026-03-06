using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\Input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\Output.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The format is automatically determined from the ".pdf" extension.
        doc.Save(outputPath);
    }
}
