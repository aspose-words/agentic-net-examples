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
        // Step 1: Create a sample document with a repeating section content control.
        Document doc = new Document();

        // Create a block‑level repeating section SDT.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeatingSection.Title = "Items";
        repeatingSection.Tag = "repeating-items";

        // Add a paragraph that will be repeated.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "First item"));
        repeatingSection.AppendChild(paragraph);

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the sample document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Step 2: Load the document and extract all repeating section content controls.
        Document loadedDoc = new Document(inputPath);

        // Find all StructuredDocumentTag nodes of type RepeatingSection.
        List<object> repeatingData = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection)
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Text = sdt.GetText().Trim()
            })
            .Cast<object>()
            .ToList();

        // Step 3: Serialize the extracted data to JSON.
        string json = JsonConvert.SerializeObject(repeatingData, Formatting.Indented);
        const string jsonPath = "repeating-sections.json";
        File.WriteAllText(jsonPath, json);

        // Optional: Save the loaded document (demonstrates a save operation).
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
