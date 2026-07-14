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
        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that contains a repeating section SDT.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a repeating section content control (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(
            doc,
            SdtType.RepeatingSection,
            MarkupLevel.Block);

        // Add some sample text inside the repeating section.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "Sample item text"));
        repeatingSection.AppendChild(paragraph);

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the sample document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract all repeating section SDTs.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Find all StructuredDocumentTag nodes of type RepeatingSection.
        List<object> extractedData = loadedDoc
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

        // -----------------------------------------------------------------
        // 3. Serialize the extracted information to JSON.
        // -----------------------------------------------------------------
        string json = JsonConvert.SerializeObject(extractedData, Formatting.Indented);
        const string jsonPath = "repeating-sections.json";
        File.WriteAllText(jsonPath, json);

        // -----------------------------------------------------------------
        // 4. (Optional) Save the loaded document again for completeness.
        // -----------------------------------------------------------------
        const string outputPath = "repeating-sections.docx";
        loadedDoc.Save(outputPath);
    }
}
