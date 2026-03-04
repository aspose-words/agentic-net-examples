using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourceFile = "MyDir/Document.docx";

        // Path where the resulting PDF will be saved.
        string destinationFile = "ArtifactsDir/Document.ConvertToPdf.pdf";

        // Load the existing DOCX document.
        Document doc = new Document(sourceFile);

        // Save the document as PDF. The format is inferred from the .pdf extension.
        doc.Save(destinationFile);
    }
}
