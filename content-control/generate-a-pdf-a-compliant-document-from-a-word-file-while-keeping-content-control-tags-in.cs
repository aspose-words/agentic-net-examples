using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX containing content controls.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph before the content controls.
        builder.Writeln("Document with content controls:");

        // Inline plain‑text content control.
        StructuredDocumentTag inlineSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        inlineSdt.Title = "CustomerName";
        inlineSdt.Tag = "customer-name";
        inlineSdt.RemoveAllChildren();
        inlineSdt.AppendChild(new Run(doc, "Contoso"));
        builder.CurrentParagraph.AppendChild(inlineSdt);
        builder.Writeln(); // Move to next line.

        // Block‑level rich‑text content control.
        StructuredDocumentTag blockSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        blockSdt.Title = "SectionContent";
        Paragraph blockParagraph = new Paragraph(doc);
        blockParagraph.AppendChild(new Run(doc, "Block content control text."));
        blockSdt.AppendChild(blockParagraph);
        doc.FirstSection.Body.AppendChild(blockSdt);

        // Save the source Word document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the Word document.
        Document loadedDoc = new Document(inputPath);

        // Configure PDF/A‑1b compliance options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF/A compliant file.
        const string outputPath = "output.pdf";
        loadedDoc.Save(outputPath, pdfOptions);
    }
}
