using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source document (can be DOCX, DOC, RTF, HTML, etc.).
        string inputPath = "input.docx";

        // Path where the PDF will be saved. The .pdf extension tells Aspose.Words to save in PDF format.
        string outputPath = "output.pdf";

        // Load the document. The constructor automatically detects the file format.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save(outputPath);
    }
}
