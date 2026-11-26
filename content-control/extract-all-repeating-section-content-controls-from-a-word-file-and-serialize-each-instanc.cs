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
        Document seedDoc = new Document();
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(seedDoc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeatingSection.Title = "Products";
        repeatingSection.Tag = "products";

        // First repeating item.
        StructuredDocumentTag item1 = new StructuredDocumentTag(seedDoc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        Paragraph p1 = new Paragraph(seedDoc);
        p1.AppendChild(new Run(seedDoc, "Product A"));
        item1.AppendChild(p1);
        repeatingSection.AppendChild(item1);

        // Second repeating item.
        StructuredDocumentTag item2 = new StructuredDocumentTag(seedDoc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        Paragraph p2 = new Paragraph(seedDoc);
        p2.AppendChild(new Run(seedDoc, "Product B"));
        item2.AppendChild(p2);
        repeatingSection.AppendChild(item2);

        // Insert the repeating section into the document body.
        seedDoc.FirstSection.Body.AppendChild(repeatingSection);
        const string inputPath = "input.docx";
        seedDoc.Save(inputPath);

        // Load the document containing repeating sections.
        Document doc = new Document(inputPath);

        // Find all repeating section content controls.
        List<object> extractedData = new List<object>();
        var repeatingSections = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                   .OfType<StructuredDocumentTag>()
                                   .Where(sdt => sdt.SdtType == SdtType.RepeatingSection);

        foreach (var section in repeatingSections)
        {
            // Get all repeating section items inside this repeating section.
            var items = section.GetChildNodes(NodeType.StructuredDocumentTag, false)
                               .OfType<StructuredDocumentTag>()
                               .Where(item => item.SdtType == SdtType.RepeatingSectionItem)
                               .ToList();

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                string text = item.GetText().Trim();
                extractedData.Add(new
                {
                    SectionTitle = section.Title,
                    SectionTag = section.Tag,
                    ItemIndex = i + 1,
                    ItemText = text
                });
            }
        }

        // Serialize the extracted information to JSON.
        string json = JsonConvert.SerializeObject(extractedData, Formatting.Indented);
        const string outputPath = "repeating-sections.json";
        File.WriteAllText(outputPath, json);

        // Optionally, save the document after processing.
        doc.Save("processed.docx");
    }
}
