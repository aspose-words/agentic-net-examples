using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

namespace MergeWordDocumentsWithContentControls
{
    public class Program
    {
        public static void Main()
        {
            // Create first sample document with a plain‑text content control.
            Document doc1 = new Document();
            Paragraph para1 = doc1.FirstSection.Body.FirstParagraph;
            para1.AppendChild(new Run(doc1, "Document 1: "));
            StructuredDocumentTag sdt1 = new StructuredDocumentTag(doc1, SdtType.PlainText, MarkupLevel.Inline);
            sdt1.Title = "Name";
            sdt1.Tag = "name";
            sdt1.RemoveAllChildren();
            sdt1.AppendChild(new Run(doc1, "Alice"));
            para1.AppendChild(sdt1);
            doc1.Save("doc1.docx");

            // Create second sample document with a plain‑text content control.
            Document doc2 = new Document();
            Paragraph para2 = doc2.FirstSection.Body.FirstParagraph;
            para2.AppendChild(new Run(doc2, "Document 2: "));
            StructuredDocumentTag sdt2 = new StructuredDocumentTag(doc2, SdtType.PlainText, MarkupLevel.Inline);
            sdt2.Title = "Address";
            sdt2.Tag = "address";
            sdt2.RemoveAllChildren();
            sdt2.AppendChild(new Run(doc2, "123 Main St"));
            para2.AppendChild(sdt2);
            doc2.Save("doc2.docx");

            // Load the first document as the base for merging.
            Document mergedDoc = new Document("doc1.docx");

            // Load the second document.
            Document secondDoc = new Document("doc2.docx");

            // Append the second document to the first one, keeping source formatting.
            mergedDoc.AppendDocument(secondDoc, ImportFormatMode.KeepSourceFormatting);

            // Note: Aspose.Words no longer provides Document.UpdateSdtIds().
            // The merge operation preserves unique IDs, so no additional action is required.

            // Save the merged document.
            mergedDoc.Save("merged.docx");

            // Extract information about all content controls in the merged document.
            List<object> sdtInfo = mergedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .Select(sdt => new
                {
                    Id = sdt.Id,
                    Title = sdt.Title,
                    Tag = sdt.Tag,
                    Text = sdt.GetText().Trim()
                })
                .Cast<object>()
                .ToList();

            // Serialize the information to JSON for verification.
            string json = JsonConvert.SerializeObject(sdtInfo, Formatting.Indented);
            File.WriteAllText("content-controls.json", json);
        }
    }
}
