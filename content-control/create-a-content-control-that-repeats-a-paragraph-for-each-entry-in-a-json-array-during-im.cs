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
        // Sample JSON array to import.
        string json = "[\"First item\",\"Second item\",\"Third item\"]";

        // Deserialize JSON into a list of strings.
        List<string> items = JsonConvert.DeserializeObject<List<string>>(json) ?? new List<string>();

        // Create a new blank document.
        Document doc = new Document();

        // Create a repeating section content control (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeatingSection.Title = "ItemsRepeatingSection";
        repeatingSection.Tag = "items-section";

        // Add a paragraph for each entry in the JSON array.
        foreach (string item in items)
        {
            Paragraph paragraph = new Paragraph(doc);
            paragraph.AppendChild(new Run(doc, item));
            repeatingSection.AppendChild(paragraph);
        }

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RepeatingSection.docx");
        doc.Save(outputPath);
    }
}
