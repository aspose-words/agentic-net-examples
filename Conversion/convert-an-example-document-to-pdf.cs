using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Example.docx";

        // Path where the PDF will be saved.
        string outputPath = @"C:\Docs\Example.pdf";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The format is determined by the .pdf extension.
        doc.Save(outputPath);
    }
}
