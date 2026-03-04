using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "MyDir/Document.docx";

        // Desired path for the resulting PDF file.
        string outputPath = "ArtifactsDir/Document.ConvertToPdf.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The format is inferred from the .pdf extension.
        doc.Save(outputPath);
    }
}
