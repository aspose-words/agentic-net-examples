using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOCX containing two content controls.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory paragraph.
        builder.Writeln("Document with content controls:");

        // ----- Inline plain‑text content control -----
        StructuredDocumentTag plainSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainSdt.RemoveAllChildren();
        plainSdt.AppendChild(new Run(doc, "Contoso"));
        Paragraph currentParagraph = builder.CurrentParagraph;
        if (currentParagraph != null)
        {
            currentParagraph.AppendChild(plainSdt);
        }
        builder.Writeln(); // Move to a new paragraph.

        // ----- Block‑level rich‑text content control -----
        StructuredDocumentTag richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "Rich text block content control."));
        richSdt.AppendChild(richParagraph);
        doc.FirstSection.Body.AppendChild(richSdt);

        // Save the source DOCX locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX for further processing.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Export information about the content controls to JSON.
        var sdtInfo = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(s => new
            {
                Title = s.Title,
                Tag = s.Tag,
                Type = s.SdtType.ToString(),
                Text = s.GetText().Trim()
            })
            .ToList();

        File.WriteAllText("content-controls.json",
            JsonConvert.SerializeObject(sdtInfo, Formatting.Indented));

        // -----------------------------------------------------------------
        // 3. Configure PDF/A‑1b save options.
        //    Use the 'Compliance' property (available in all supported versions)
        //    to request PDF/A conformance.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Request PDF/A‑1b compliance.
            Compliance = PdfCompliance.PdfA1b,
            // Preserve the document structure so that content controls become PDF tags.
            ExportDocumentStructure = true
        };

        // Save the document as PDF/A.
        const string outputPdf = "output.pdf";
        loadedDoc.Save(outputPdf, pdfOptions);
    }
}
