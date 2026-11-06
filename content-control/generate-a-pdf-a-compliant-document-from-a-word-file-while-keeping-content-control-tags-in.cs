using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a plain‑text inline content control.
        builder.Writeln("Customer:");
        StructuredDocumentTag plainSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        plainSdt.Title = "CustomerName";
        plainSdt.Tag = "customer-name";
        plainSdt.RemoveAllChildren();
        plainSdt.AppendChild(new Run(doc, "Contoso"));
        // Append the content control to the current paragraph.
        builder.CurrentParagraph?.AppendChild(plainSdt);

        // Add a block‑level rich‑text content control.
        builder.Writeln(); // start a new paragraph.
        StructuredDocumentTag richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richSdt.Title = "SectionContent";
        Paragraph innerParagraph = new Paragraph(doc);
        innerParagraph.AppendChild(new Run(doc, "This is a block content control."));
        richSdt.AppendChild(innerParagraph);
        // Insert the block content control into the document body.
        doc.FirstSection.Body.AppendChild(richSdt);

        // Save the source DOCX file.
        const string sourceDocx = "sample.docx";
        doc.Save(sourceDocx);

        // Load the DOCX file (simulating an existing file).
        Document loadedDoc = new Document(sourceDocx);

        // Configure PDF/A‑1b compliance options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF/A compliant file.
        const string outputPdf = "output-pdfa.pdf";
        loadedDoc.Save(outputPdf, pdfOptions);
    }
}
