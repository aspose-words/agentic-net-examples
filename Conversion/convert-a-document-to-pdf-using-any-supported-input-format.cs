using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source document (can be DOCX, RTF, HTML, etc.).
        string inputPath = "input.docx";

        // Path where the PDF will be saved.
        string outputPath = "output.pdf";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Save the loaded document as PDF. The format is inferred from the .pdf extension.
        doc.Save(outputPath);
    }
}
