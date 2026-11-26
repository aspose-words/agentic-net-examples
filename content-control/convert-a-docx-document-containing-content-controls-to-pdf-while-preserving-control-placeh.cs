using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with content controls.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Document with content controls:");

        // Inline plain‑text content control with placeholder text.
        var plainSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        // Add placeholder text inside the SDT.
        plainSdt.RemoveAllChildren();
        plainSdt.AppendChild(new Run(doc, "Enter customer name"));
        // Append the SDT to the current paragraph.
        builder.CurrentParagraph.AppendChild(plainSdt);
        builder.Writeln(); // Move to next line.

        // Block‑level rich‑text content control with placeholder text.
        var richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Comments",
            Tag = "comments"
        };
        // Add a paragraph containing placeholder text.
        var innerParagraph = new Paragraph(doc);
        innerParagraph.AppendChild(new Run(doc, "Enter comments"));
        richSdt.AppendChild(innerParagraph);
        // Append the block SDT to the document body.
        doc.FirstSection.Body.AppendChild(richSdt);

        // Save the sample DOCX.
        const string docxPath = "sample.docx";
        doc.Save(docxPath);

        // Load the DOCX and convert it to PDF, preserving placeholders.
        var loadedDoc = new Document(docxPath);
        const string pdfPath = "sample.pdf";
        loadedDoc.Save(pdfPath);

        // Extract metadata of all content controls.
        var sdtInfos = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .Cast<StructuredDocumentTag>()
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Type = sdt.SdtType.ToString(),
                Placeholder = sdt.PlaceholderName, // Will be empty because we used inline text.
                Text = sdt.GetText().Trim()
            })
            .ToList();

        // Serialize metadata to JSON.
        const string jsonPath = "sdt-info.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sdtInfos, Formatting.Indented));
    }
}
