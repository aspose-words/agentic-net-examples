using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document with a plain‑text content control.
        Document seedDoc = new Document();
        StructuredDocumentTag sdt = new StructuredDocumentTag(seedDoc, SdtType.PlainText, MarkupLevel.Inline);
        sdt.Title = "SampleControl";
        sdt.Tag = "sample-control";
        sdt.AppendChild(new Run(seedDoc, "Original content"));
        Paragraph para = seedDoc.FirstSection.Body.FirstParagraph;
        para.AppendChild(sdt);
        const string inputPath = "input.docx";
        seedDoc.Save(inputPath);

        // Step 2: Load the document that contains the content control.
        Document doc = new Document(inputPath);

        // Step 3: Locate the content control by its Title.
        StructuredDocumentTag targetSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(tag => tag.Title == "SampleControl");

        if (targetSdt == null)
        {
            throw new InvalidOperationException("Content control with the specified title was not found.");
        }

        // Step 4: Clear the contents of the content control without removing the control itself.
        targetSdt.RemoveAllChildren();

        // Step 5: Save the modified document.
        const string outputPath = "cleared.docx";
        doc.Save(outputPath);
    }
}
