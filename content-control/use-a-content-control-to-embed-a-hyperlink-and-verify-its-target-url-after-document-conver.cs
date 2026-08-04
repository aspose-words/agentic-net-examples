using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "LinkControl",
            Tag = "link-sdt"
        };

        // The content control must contain at least one paragraph.
        Paragraph sdtParagraph = new Paragraph(doc);
        sdt.AppendChild(sdtParagraph);

        // Insert the hyperlink inside the paragraph that belongs to the content control.
        builder.MoveTo(sdtParagraph);
        builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);

        // Add the content control to the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Save the original DOCX.
        const string docxPath = "linkControl.docx";
        doc.Save(docxPath);

        // Convert the document to PDF.
        const string pdfPath = "linkControl.pdf";
        doc.Save(pdfPath);

        // Load the PDF back into a Document object.
        Document pdfDoc = new Document(pdfPath);

        // Locate the first hyperlink field in the converted document.
        FieldHyperlink hyperlink = pdfDoc.Range.Fields
            .OfType<FieldHyperlink>()
            .FirstOrDefault();

        // Verify that the hyperlink's target URL is the expected one.
        string address = hyperlink?.Address ?? "NotFound";
        bool isCorrect = address == "https://www.aspose.com";

        // Output the verification result.
        Console.WriteLine($"Hyperlink address after conversion: {address}");
        Console.WriteLine($"Verification: {(isCorrect ? "Success" : "Failure")}");
    }
}
