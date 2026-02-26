using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderToPdfExample
{
    static void Main()
    {
        // Define input and output locations.
        // Adjust these paths to point to your actual files/folders.
        string inputFile = @"C:\Docs\SampleDocument.docx";
        string outputFile = @"C:\Docs\SampleDocument.pdf";

        // Load the existing Word document.
        Document doc = new Document(inputFile); // uses Document(string) constructor

        // Optional: if you need to customize PDF conversion, create PdfSaveOptions.
        // Here we use the default options, but the object is created via the rule.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF using the Save(string, SaveOptions) overload.
        doc.Save(outputFile, pdfOptions); // saves as PDF because of .pdf extension
    }
}
