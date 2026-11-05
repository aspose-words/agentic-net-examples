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
        // Create a sample document with several content controls.
        Document seedDoc = new Document();
        Body body = seedDoc.FirstSection.Body;

        // First content control: plain text, inline.
        StructuredDocumentTag sdt1 = new StructuredDocumentTag(seedDoc, SdtType.PlainText, MarkupLevel.Inline);
        sdt1.Title = "Customer Name";
        sdt1.Tag = "old-tag-1";
        sdt1.RemoveAllChildren();
        sdt1.AppendChild(new Run(seedDoc, "Alice"));
        body.FirstParagraph.AppendChild(sdt1);

        // Second content control: rich text, block.
        StructuredDocumentTag sdt2 = new StructuredDocumentTag(seedDoc, SdtType.RichText, MarkupLevel.Block);
        sdt2.Title = "Address";
        sdt2.Tag = "old-tag-2";
        Paragraph para2 = new Paragraph(seedDoc);
        para2.AppendChild(new Run(seedDoc, "123 Main St"));
        sdt2.AppendChild(para2);
        body.AppendChild(sdt2);

        // Third content control: dropdown list, inline.
        StructuredDocumentTag sdt3 = new StructuredDocumentTag(seedDoc, SdtType.DropDownList, MarkupLevel.Inline);
        sdt3.Title = "Country";
        sdt3.Tag = "old-tag-3";
        sdt3.ListItems.Add(new SdtListItem("USA", "US"));
        sdt3.ListItems.Add(new SdtListItem("Canada", "CA"));
        body.FirstParagraph.AppendChild(sdt3);

        // Save the seed document.
        const string inputPath = "input.docx";
        seedDoc.Save(inputPath);

        // Load the document for processing.
        Document doc = new Document(inputPath);

        // Enumerate all StructuredDocumentTag nodes.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                          .OfType<StructuredDocumentTag>()
                          .ToList();

        // Update each content control's Tag to a standardized format.
        int index = 1;
        foreach (var sdt in sdtNodes)
        {
            // Use the Title if available; otherwise fall back to an index-based tag.
            string baseTag = !string.IsNullOrEmpty(sdt.Title)
                ? sdt.Title.Trim().ToLowerInvariant().Replace(' ', '_')
                : $"sdt_{index}";

            // Ensure uniqueness by appending the index.
            string standardizedTag = $"{baseTag}_{index}";
            sdt.Tag = standardizedTag;
            index++;
        }

        // Save the updated document.
        const string outputPath = "updated.docx";
        doc.Save(outputPath);
    }
}
