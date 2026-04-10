using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Prepare sample source documents with content controls.
        string doc1Path = "Doc1.docx";
        string doc2Path = "Doc2.docx";

        CreateSampleDocument(doc1Path, "First document content control", "FirstControl");
        CreateSampleDocument(doc2Path, "Second document content control", "SecondControl");

        // Load the source documents.
        Document srcDoc1 = new Document(doc1Path);
        Document srcDoc2 = new Document(doc2Path);

        // Create the destination document.
        Document mergedDoc = new Document();

        // Append the first source document.
        mergedDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);

        // Append the second source document.
        mergedDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Update IDs (CustomNodeId) of all content controls in the merged document.
        NodeCollection sdtNodes = mergedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // CustomNodeId expects an integer in this version of Aspose.Words.
            sdt.CustomNodeId = Guid.NewGuid().GetHashCode();
        }

        // Save the merged document.
        string mergedPath = "Merged.docx";
        mergedDoc.Save(mergedPath);
    }

    // Helper method to create a document containing a single plain‑text content control.
    private static void CreateSampleDocument(string filePath, string controlText, string controlTag)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph before the content control.
        builder.Writeln("Paragraph before the content control.");

        // Create an inline plain‑text StructuredDocumentTag.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleControl",
            Tag = controlTag
        };

        // Append the SDT to the current paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(sdt);

        // Add text inside the content control.
        sdt.AppendChild(new Run(doc, controlText));

        // Insert another paragraph after the control.
        builder.Writeln("Paragraph after the content control.");

        // Save the document.
        doc.Save(filePath);
    }
}
