using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class ExportContentControlsToXps
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -------------------------
        // Insert an inline plain‑text content control.
        // -------------------------
        Paragraph para1 = doc.FirstSection.Body.FirstParagraph;
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "Contoso"));
        para1.AppendChild(plainTextSdt);

        // -------------------------
        // Insert a block‑level rich‑text content control.
        // -------------------------
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "SectionContent",
            Tag = "section-content"
        };
        Paragraph richPara = new Paragraph(doc);
        richPara.AppendChild(new Run(doc, "This is a block‑level rich text content control."));
        richTextSdt.AppendChild(richPara);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // -------------------------
        // Insert an inline checkbox content control.
        // -------------------------
        Paragraph para2 = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para2);
        StructuredDocumentTag checkboxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AgreeTerms",
            Tag = "agree-terms",
            Checked = true
        };
        para2.AppendChild(checkboxSdt);

        // -------------------------
        // Save the document to XPS format while preserving content control boundaries.
        // -------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ContentControlsOutput.xps");
        XpsSaveOptions xpsOptions = new XpsSaveOptions(SaveFormat.Xps);
        // The default behavior preserves content control boundaries; no extra options required.
        doc.Save(outputPath, xpsOptions);
    }
}
