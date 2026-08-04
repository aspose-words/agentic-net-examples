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
        // Create two sample source documents that contain a single plain‑text content control.
        Document sourceDoc1 = CreateSampleDocument("FirstDocument", "FirstTag", "First value");
        Document sourceDoc2 = CreateSampleDocument("SecondDocument", "SecondTag", "Second value");

        // Save the source documents to the local file system (required for the load step).
        sourceDoc1.Save("Source1.docx");
        sourceDoc2.Save("Source2.docx");

        // Load the source documents (simulating a real‑world scenario where files already exist).
        Document src1 = new Document("Source1.docx");
        Document src2 = new Document("Source2.docx");

        // Create the destination document that will receive the merged content.
        Document destination = new Document();

        // Append the source documents while keeping their original formatting.
        destination.AppendDocument(src1, ImportFormatMode.KeepSourceFormatting);
        destination.AppendDocument(src2, ImportFormatMode.KeepSourceFormatting);

        // Update the custom IDs of all content controls in the merged document.
        NodeCollection sdtNodes = destination.GetChildNodes(NodeType.StructuredDocumentTag, true);
        int idCounter = 1;
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // In Aspose.Words v2 the CustomNodeId property is an integer.
            sdt.CustomNodeId = idCounter++;
        }

        // Export information about the content controls to a JSON file (optional reporting).
        var controlInfo = destination.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(s => new
            {
                Title = s.Title,
                Tag = s.Tag,
                CustomNodeId = s.CustomNodeId
            })
            .ToList();

        string json = JsonConvert.SerializeObject(controlInfo, Formatting.Indented);
        File.WriteAllText("ContentControlsInfo.json", json);

        // Save the merged document.
        destination.Save("MergedDocument.docx");
    }

    // Helper method that creates a document containing a single plain‑text content control.
    private static Document CreateSampleDocument(string title, string tag, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading so each source document is identifiable.
        builder.Writeln($"--- {title} ---");

        // Create an inline plain‑text StructuredDocumentTag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = title,
            Tag = tag
        };
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, text));

        // Insert the content control into the current paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(sdt);

        return doc;
    }
}
