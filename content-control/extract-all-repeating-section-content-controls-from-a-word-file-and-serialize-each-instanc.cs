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
        // Create a sample document with repeating section content controls.
        var sourceDoc = new Document();

        // First repeating section.
        var repeating1 = new StructuredDocumentTag(sourceDoc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "RepeatingSection1",
            Tag = "rep1"
        };
        // First item of the first repeating section.
        var item1a = new StructuredDocumentTag(sourceDoc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        var para1a = new Paragraph(sourceDoc);
        para1a.AppendChild(new Run(sourceDoc, "Item 1A"));
        item1a.AppendChild(para1a);
        repeating1.AppendChild(item1a);
        // Second item of the first repeating section.
        var item1b = new StructuredDocumentTag(sourceDoc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        var para1b = new Paragraph(sourceDoc);
        para1b.AppendChild(new Run(sourceDoc, "Item 1B"));
        item1b.AppendChild(para1b);
        repeating1.AppendChild(item1b);
        sourceDoc.FirstSection.Body.AppendChild(repeating1);

        // Second repeating section.
        var repeating2 = new StructuredDocumentTag(sourceDoc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "RepeatingSection2",
            Tag = "rep2"
        };
        // Single item of the second repeating section.
        var item2a = new StructuredDocumentTag(sourceDoc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        var para2a = new Paragraph(sourceDoc);
        para2a.AppendChild(new Run(sourceDoc, "Item 2A"));
        item2a.AppendChild(para2a);
        repeating2.AppendChild(item2a);
        sourceDoc.FirstSection.Body.AppendChild(repeating2);

        // Save the sample document.
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath);

        // Load the document for processing.
        var doc = new Document(inputPath);

        // Find all repeating section content controls.
        var repeatingControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection)
            .ToList();

        // Prepare data for JSON serialization.
        var payload = new List<object>();
        foreach (var repeating in repeatingControls)
        {
            var items = repeating.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .Where(item => item.SdtType == SdtType.RepeatingSectionItem)
                .Select(item => item.GetText().Trim())
                .ToList();

            payload.Add(new
            {
                Title = repeating.Title ?? string.Empty,
                Tag = repeating.Tag ?? string.Empty,
                Items = items
            });
        }

        // Serialize to JSON.
        string json = JsonConvert.SerializeObject(payload, Formatting.Indented);
        const string jsonPath = "repeating-sections.json";
        File.WriteAllText(jsonPath, json);
    }
}
