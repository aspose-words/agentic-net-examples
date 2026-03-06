using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "MyDir/Document.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "ArtifactsDir/Document.ConvertToPdf.pdf";

        // Load the DOCX document using the Document(string) constructor.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The Save(string) method infers the format from the file extension.
        doc.Save(outputPath);
    }
}
