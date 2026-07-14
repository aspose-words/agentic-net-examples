using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with various content controls.
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);

        // Plain text content control.
        var plainTextSdt = new StructuredDocumentTag(sourceDoc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Customer Name",
            Tag = "custName"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(sourceDoc, "John Doe"));
        builder.CurrentParagraph.AppendChild(plainTextSdt);
        builder.Writeln();

        // Rich text content control.
        var richTextSdt = new StructuredDocumentTag(sourceDoc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Address Block",
            Tag = "addrBlock"
        };
        var para = new Paragraph(sourceDoc);
        para.AppendChild(new Run(sourceDoc, "123 Main St"));
        richTextSdt.AppendChild(para);
        sourceDoc.FirstSection.Body.AppendChild(richTextSdt);
        builder.Writeln();

        // Drop-down list content control.
        var dropdownSdt = new StructuredDocumentTag(sourceDoc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Country Selector",
            Tag = "countrySel"
        };
        dropdownSdt.ListItems.Add(new SdtListItem("USA", "US"));
        dropdownSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
        builder.CurrentParagraph.AppendChild(dropdownSdt);
        builder.Writeln();

        // Save the source document.
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath);

        // Load the document for processing.
        var doc = new Document(inputPath);

        // Prepare a list to capture tag changes for reporting.
        var tagChanges = new List<object>();

        // Enumerate all StructuredDocumentTag nodes and update their Tag.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>();

        foreach (var sdt in sdtNodes)
        {
            var oldTag = sdt.Tag ?? string.Empty;
            var title = sdt.Title ?? "untitled";

            // Standardized naming: lower‑case title with hyphens, prefixed with "sdt-".
            var standardizedTag = "sdt-" + string.Concat(title
                .Trim()
                .Split(new[] { ' ', '_' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(part => part.ToLowerInvariant()))
                .Replace(" ", "-");

            sdt.Tag = standardizedTag;

            tagChanges.Add(new
            {
                Title = title,
                OldTag = oldTag,
                NewTag = standardizedTag
            });
        }

        // Serialize the tag change report to JSON.
        const string jsonPath = "tags-updated.json";
        var json = JsonConvert.SerializeObject(tagChanges, Formatting.Indented);
        File.WriteAllText(jsonPath, json);

        // Save the updated document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
