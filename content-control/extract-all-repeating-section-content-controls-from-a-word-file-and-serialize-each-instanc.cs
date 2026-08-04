using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two repeating section content controls.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First repeating section.
        StructuredDocumentTag repeating1 = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "FirstSection",
            Tag = "first"
        };
        Paragraph para1 = new Paragraph(doc);
        para1.AppendChild(new Run(doc, "First item content"));
        repeating1.AppendChild(para1);
        doc.FirstSection.Body.AppendChild(repeating1);

        // Second repeating section.
        StructuredDocumentTag repeating2 = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "SecondSection",
            Tag = "second"
        };
        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "Second item content"));
        repeating2.AppendChild(para2);
        doc.FirstSection.Body.AppendChild(repeating2);

        // Save the sample document.
        const string samplePath = "sample.docx";
        doc.Save(samplePath);

        // Load the document from file.
        Document loadedDoc = new Document(samplePath);

        // Extract all repeating section content controls.
        var repeatingControls = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection)
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Text = sdt.GetText().Trim()
            })
            .ToList();

        // Serialize the extracted data to JSON.
        string json = JsonConvert.SerializeObject(repeatingControls, Formatting.Indented);
        File.WriteAllText("repeating-sections.json", json);

        // Optionally, save the loaded document to demonstrate a full lifecycle.
        loadedDoc.Save("output.docx");
    }
}
