using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        const string docPath = "sample.docx";
        const string jsonPath = "repeating_sections.json";

        // Create a sample document with a repeating section content control.
        CreateSampleDocument(docPath);

        // Load the document.
        Document doc = new Document(docPath);

        // Extract repeating section items.
        List<RepeatingSectionItemData> items = ExtractRepeatingSectionItems(doc);

        // Serialize to JSON.
        string json = JsonConvert.SerializeObject(items, Formatting.Indented);
        File.WriteAllText(jsonPath, json);
    }

    private static void CreateSampleDocument(string path)
    {
        // Create a new blank document.
        Document doc = new Document();
        Body body = doc.FirstSection.Body;

        // Insert a block‑level repeating section content control into the document body.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        body.AppendChild(repeatingSection);

        // Add three repeating section items.
        for (int i = 1; i <= 3; i++)
        {
            // Create a repeating section item (also block‑level).
            StructuredDocumentTag item = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Block);

            // Each item must contain at least one paragraph with text.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, $"Item {i} text");
            para.AppendChild(run);
            item.AppendChild(para);

            // Append the item to the repeating section.
            repeatingSection.AppendChild(item);
        }

        // Save the document.
        doc.Save(path);
    }

    private static List<RepeatingSectionItemData> ExtractRepeatingSectionItems(Document doc)
    {
        List<RepeatingSectionItemData> result = new List<RepeatingSectionItemData>();

        // Get all StructuredDocumentTag nodes in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            if (sdt.SdtType == SdtType.RepeatingSection)
            {
                // Find all repeating section items that belong to this repeating section.
                NodeCollection itemNodes = sdt.GetChildNodes(NodeType.StructuredDocumentTag, true);
                int index = 0;

                foreach (StructuredDocumentTag item in itemNodes)
                {
                    if (item.SdtType == SdtType.RepeatingSectionItem)
                    {
                        string content = item.GetText().Trim();
                        result.Add(new RepeatingSectionItemData
                        {
                            Index = index,
                            Content = content
                        });
                        index++;
                    }
                }
            }
        }

        return result;
    }

    private class RepeatingSectionItemData
    {
        public int Index { get; set; }
        public string Content { get; set; } = string.Empty;
    }
}
