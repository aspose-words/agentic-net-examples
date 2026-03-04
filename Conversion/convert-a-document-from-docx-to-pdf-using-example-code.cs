using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Input folder containing the DOCX file
        string MyDir = @"C:\Input\";
        // Output folder where the PDF will be saved
        string ArtifactsDir = @"C:\Output\";

        // Load the existing DOCX document
        Document doc = new Document(MyDir + "Document.docx");

        // Save the document as PDF; the format is inferred from the .pdf extension
        doc.Save(ArtifactsDir + "Document.ConvertToPdf.pdf");
    }
}
