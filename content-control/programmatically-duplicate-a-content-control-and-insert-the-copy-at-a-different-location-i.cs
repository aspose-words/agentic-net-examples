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
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to host the original content control.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph ?? new Paragraph(doc);
        if (doc.FirstSection.Body.FirstParagraph == null)
        {
            doc.FirstSection.Body.AppendChild(firstParagraph);
        }

        // Create an inline plain‑text content control.
        StructuredDocumentTag originalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "OriginalControl",
            Tag = "original"
        };
        originalSdt.RemoveAllChildren();
        originalSdt.AppendChild(new Run(doc, "Original Content"));
        firstParagraph.AppendChild(originalSdt);

        // Add a second paragraph where the duplicate will be placed.
        Paragraph secondParagraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(secondParagraph);

        // Clone the original content control (deep clone) and adjust its metadata.
        StructuredDocumentTag clonedSdt = (StructuredDocumentTag)originalSdt.Clone(true);
        clonedSdt.Title = "ClonedControl";
        clonedSdt.Tag = "cloned";

        // Insert the cloned content control into the second paragraph.
        secondParagraph.AppendChild(clonedSdt);

        // Save the resulting document.
        const string outputDocPath = "DuplicatedContentControl.docx";
        doc.Save(outputDocPath);

        // Export information about all content controls to JSON.
        var controlsInfo = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Text = sdt.GetText().Trim()
            })
            .ToList();

        string json = JsonConvert.SerializeObject(controlsInfo, Formatting.Indented);
        File.WriteAllText("contentControls.json", json);
    }
}
