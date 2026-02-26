using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "MyDir/Document.docx";

        // Path where the PDF will be saved. The .pdf extension tells Aspose.Words to save in PDF format.
        string outputPath = "ArtifactsDir/Document.ConvertToPdf.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The format is automatically determined from the file extension.
        doc.Save(outputPath);
    }
}
