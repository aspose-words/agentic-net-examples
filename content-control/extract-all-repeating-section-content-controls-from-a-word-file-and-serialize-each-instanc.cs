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
        // Create a sample document that contains a repeating section content control.
        const string inputPath = "input.docx";
        Document doc = new Document();

        // Create a repeating section SDT (block level) and add a paragraph with sample text.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeatingSection.Title = "SampleRepeatingSection";
        repeatingSection.Tag = "sample-repeating";

        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "Item 1"));
        repeatingSection.AppendChild(paragraph);

        // Append the repeating section to the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the document so it can be loaded later.
        doc.Save(inputPath);

        // Load the document that contains the repeating section controls.
        Document loadedDoc = new Document(inputPath);

        // Find all repeating section content controls in the document.
        List<StructuredDocumentTag> repeatingControls = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection)
            .ToList();

        // Prepare a serializable model for each repeating section instance.
        var payload = repeatingControls
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Text = sdt.GetText().Trim()
            })
            .ToList();

        // Serialize the collection to JSON with indentation.
        string json = JsonConvert.SerializeObject(payload, Formatting.Indented);

        // Write the JSON to a file.
        const string jsonPath = "repeating-sections.json";
        File.WriteAllText(jsonPath, json);

        // Optionally, save the loaded document (unchanged) to demonstrate a second output file.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
