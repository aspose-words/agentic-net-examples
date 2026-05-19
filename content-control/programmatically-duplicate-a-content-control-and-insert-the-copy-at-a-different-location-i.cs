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
        // Create a source document with a block-level plain‑text content control.
        Document sourceDoc = new Document();
        StructuredDocumentTag originalSdt = new StructuredDocumentTag(sourceDoc, SdtType.PlainText, MarkupLevel.Block);
        originalSdt.Title = "SampleControl";
        originalSdt.Tag = "sample-control";

        // Add some text inside the content control.
        Paragraph sdtParagraph = new Paragraph(sourceDoc);
        sdtParagraph.AppendChild(new Run(sourceDoc, "Original content"));
        originalSdt.AppendChild(sdtParagraph);

        // Insert the content control into the document body.
        sourceDoc.FirstSection.Body.AppendChild(originalSdt);

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document back.
        Document doc = new Document(sourcePath);

        // Locate the original content control by its Title.
        StructuredDocumentTag foundSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(s => s.Title == "SampleControl");

        if (foundSdt == null)
            throw new InvalidOperationException("Original content control not found.");

        // Clone the content control (deep clone, including its children).
        StructuredDocumentTag clonedSdt = (StructuredDocumentTag)foundSdt.Clone(true);

        // Optionally modify the cloned control (e.g., change its title to avoid duplicates).
        clonedSdt.Title = "SampleControlCopy";
        clonedSdt.Tag = "sample-control-copy";

        // Insert the cloned control after the original one.
        doc.FirstSection.Body.InsertAfter(clonedSdt, foundSdt);

        // Save the resulting document with the duplicated content control.
        const string resultPath = "duplicated.docx";
        doc.Save(resultPath);
    }
}
