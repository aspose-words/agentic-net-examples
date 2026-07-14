using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json.Linq;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a block‑level repeating section content control.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Create a template item inside the repeating section.
        StructuredDocumentTag itemTemplate = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        repeatingSection.AppendChild(itemTemplate);

        // Add a paragraph with placeholder text to the template item.
        Paragraph templateParagraph = new Paragraph(doc);
        templateParagraph.AppendChild(new Run(doc, "Placeholder"));
        itemTemplate.AppendChild(templateParagraph);

        // Sample JSON array to import.
        string json = @"[
            { ""Name"": ""Alice"" },
            { ""Name"": ""Bob"" },
            { ""Name"": ""Charlie"" }
        ]";

        // Parse the JSON array.
        JArray dataArray = JArray.Parse(json);

        // For each JSON object, clone the template item and set its paragraph text.
        foreach (JObject obj in dataArray)
        {
            // Clone the template (deep clone).
            StructuredDocumentTag clonedItem = (StructuredDocumentTag)itemTemplate.Clone(true);

            // Retrieve the paragraph inside the cloned item.
            Paragraph clonedParagraph = (Paragraph)clonedItem.GetChild(NodeType.Paragraph, 0, true);
            if (clonedParagraph != null && clonedParagraph.Runs.Count > 0)
            {
                // Replace the placeholder text with the value from JSON.
                clonedParagraph.Runs[0].Text = obj["Name"]?.ToString() ?? string.Empty;
            }

            // Append the populated item to the repeating section.
            repeatingSection.AppendChild(clonedItem);
        }

        // Remove the original template item; it is no longer needed.
        itemTemplate.Remove();

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RepeatingSectionFromJson.docx");
        doc.Save(outputPath);
    }
}
