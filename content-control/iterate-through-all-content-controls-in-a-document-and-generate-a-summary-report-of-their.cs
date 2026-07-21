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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Plain‑text block content control ----------
        // Create a block‑level SDT, add a paragraph with a run, then append it to the document body.
        StructuredDocumentTag plain = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        plain.Title = "PlainText";
        plain.Tag = "plain";

        Paragraph plainPara = new Paragraph(doc);
        plainPara.AppendChild(new Run(doc, "Plain text content"));
        plain.AppendChild(plainPara);

        doc.FirstSection.Body.AppendChild(plain);
        builder.Writeln(); // Move to a new paragraph.

        // ---------- Rich‑text block content control ----------
        StructuredDocumentTag rich = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        rich.Title = "RichText";
        rich.Tag = "rich";

        Paragraph richPara = new Paragraph(doc);
        richPara.AppendChild(new Run(doc, "Rich text content"));
        rich.AppendChild(richPara);

        doc.FirstSection.Body.AppendChild(rich);
        builder.Writeln();

        // ---------- Checkbox inline content control ----------
        // Move to a new paragraph and insert an inline checkbox SDT.
        builder.Writeln();
        StructuredDocumentTag checkBox = builder.InsertStructuredDocumentTag(SdtType.Checkbox);
        checkBox.Title = "Agree";
        checkBox.Tag = "agree";
        checkBox.Checked = true;
        builder.Writeln();

        // ---------- Drop‑down list inline content control ----------
        builder.Writeln();
        StructuredDocumentTag dropDown = builder.InsertStructuredDocumentTag(SdtType.DropDownList);
        dropDown.Title = "Options";
        dropDown.Tag = "options";
        dropDown.ListItems.Add(new SdtListItem("Option 1", "1"));
        dropDown.ListItems.Add(new SdtListItem("Option 2", "2"));
        builder.Writeln();

        // Save the sample document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Gather information about all content controls in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        var reportItems = sdtNodes
            .OfType<StructuredDocumentTag>()
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Type = sdt.SdtType.ToString()
            })
            .Cast<object>()
            .ToList();

        // Serialize the report to JSON.
        string jsonReport = JsonConvert.SerializeObject(reportItems, Formatting.Indented);
        const string jsonPath = "content-controls-report.json";
        File.WriteAllText(jsonPath, jsonReport);

        // Output paths for verification.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
        Console.WriteLine($"Report saved to: {Path.GetFullPath(jsonPath)}");
    }
}
