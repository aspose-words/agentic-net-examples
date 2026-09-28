using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // ---------- Inline plain‑text content control ----------
        Paragraph para1 = doc.FirstSection.Body.FirstParagraph;
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "Contoso"));
        para1.AppendChild(plainTextSdt);

        // ---------- Inline checkbox content control ----------
        StructuredDocumentTag checkboxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AgreeTerms",
            Tag = "agree-terms",
            Checked = true
        };
        para1.AppendChild(new Run(doc, " "));
        para1.AppendChild(checkboxSdt);

        // ---------- Block‑level rich‑text content control ----------
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "SectionContent",
            Tag = "section-content"
        };
        Paragraph blockPara = new Paragraph(doc);
        blockPara.AppendChild(new Run(doc, "This is a block‑level rich text content control."));
        richTextSdt.AppendChild(blockPara);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the DOCX for reference (optional).
        doc.Save("ContentControls.docx");

        // Export to XPS while preserving content control boundaries.
        // XpsSaveOptions does not expose ExportDocumentStructure; the default behavior retains the structure.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        doc.Save("ContentControls.xps", xpsOptions);

        // Indicate completion.
        Console.WriteLine("Document with content controls saved as DOCX and exported to XPS.");
    }
}
