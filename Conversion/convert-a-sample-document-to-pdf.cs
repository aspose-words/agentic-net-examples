using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = "MyDir\\Document.docx";

        // Path where the PDF will be saved. The .pdf extension tells Aspose.Words to save in PDF format.
        string outputPath = "ArtifactsDir\\Document.ConvertToPdf.pdf";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Save the loaded document as PDF.
        doc.Save(outputPath);
    }
}
