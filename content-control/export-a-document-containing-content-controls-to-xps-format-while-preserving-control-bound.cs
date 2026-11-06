using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph of the document (created by default).
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // ----- Inline plain text content control -----
        StructuredDocumentTag plainSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        plainSdt.Title = "CustomerName";
        plainSdt.Tag = "customer-name";
        plainSdt.RemoveAllChildren();
        plainSdt.AppendChild(new Run(doc, "Contoso"));
        firstParagraph.AppendChild(plainSdt);

        // ----- Block-level rich text content control -----
        StructuredDocumentTag richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "This is a block-level rich text content control."));
        richSdt.AppendChild(richParagraph);
        doc.FirstSection.Body.AppendChild(richSdt);

        // ----- Inline checkbox content control -----
        StructuredDocumentTag checkSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
        checkSdt.Title = "AcceptTerms";
        checkSdt.Checked = true;
        // Add a space before the checkbox for readability.
        firstParagraph.AppendChild(new Run(doc, " "));
        firstParagraph.AppendChild(checkSdt);
        firstParagraph.AppendChild(new Run(doc, " I accept the terms."));

        // Save the document as DOCX (optional verification).
        doc.Save("ContentControls.docx");

        // Export the document to XPS format while preserving content control boundaries.
        doc.Save("ContentControls.xps", SaveFormat.Xps);

        // Serialize basic information about the content controls to JSON.
        var sdtInfo = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(s => new
            {
                Title = s.Title,
                Tag = s.Tag,
                Type = s.SdtType.ToString(),
                Text = s.GetText().Trim()
            })
            .ToList();

        File.WriteAllText("ContentControlsInfo.json",
            JsonConvert.SerializeObject(sdtInfo, Formatting.Indented));
    }
}
