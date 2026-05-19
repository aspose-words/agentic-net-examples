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
        // Create a sample document with a repeating section content control.
        Document doc = new Document();

        // Create the repeating section SDT (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "MyRepeatingSection",
            Tag = "my-repeating-section"
        };

        // First repeating item.
        StructuredDocumentTag item1 = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        Paragraph para1 = new Paragraph(doc);
        para1.AppendChild(new Run(doc, "First item"));
        item1.AppendChild(para1);
        repeatingSection.AppendChild(item1);

        // Second repeating item (clone of the first with different text).
        StructuredDocumentTag item2 = (StructuredDocumentTag)item1.Clone(true);
        item2.RemoveAllChildren();
        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "Second item"));
        item2.AppendChild(para2);
        repeatingSection.AppendChild(item2);

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the sample document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Find all repeating section content controls.
        List<StructuredDocumentTag> repeatingControls = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection)
            .ToList();

        // Prepare a collection to hold extracted data.
        var extractedItems = new List<object>();

        foreach (var repeating in repeatingControls)
        {
            // Find all repeating section items inside the current repeating section.
            List<StructuredDocumentTag> items = repeating
                .GetChildNodes(NodeType.StructuredDocumentTag, false)
                .OfType<StructuredDocumentTag>()
                .Where(sdt => sdt.SdtType == SdtType.RepeatingSectionItem)
                .ToList();

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                extractedItems.Add(new
                {
                    SectionTitle = repeating.Title,
                    SectionTag = repeating.Tag,
                    ItemIndex = i + 1,
                    Text = item.GetText().Trim()
                });
            }
        }

        // Serialize the extracted data to JSON.
        string json = JsonConvert.SerializeObject(extractedItems, Formatting.Indented);
        const string jsonPath = "repeating-sections.json";
        File.WriteAllText(jsonPath, json);

        // Optionally, save the processed document (unchanged in this example).
        const string outputDocPath = "repeating-sections.docx";
        loadedDoc.Save(outputDocPath);
    }
}
