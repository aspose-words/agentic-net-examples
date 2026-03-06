using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Set the desired page size for the first section.
        // Example: use the predefined A4 size.
        doc.FirstSection.PageSetup.PaperSize = PaperSize.A4;

        // If a custom size is required, uncomment and set the dimensions in points.
        // doc.FirstSection.PageSetup.PageWidth = 600;   // Width in points
        // doc.FirstSection.PageSetup.PageHeight = 800;  // Height in points

        // Rebuild the page layout so that the changes take effect.
        doc.UpdatePageLayout();

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save("output.pdf");
    }
}
