using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph of the document body (it always exists in a new document).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create an inline plain‑text content control.
        StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        contentControl.Title = "ReadOnlyInfo";
        contentControl.Tag = "readonly-info";

        // Lock the content control so the user cannot edit its contents or its properties.
        contentControl.LockContents = true;
        contentControl.LockContentControl = true;

        // Set the displayed text inside the content control.
        contentControl.RemoveAllChildren();
        contentControl.AppendChild(new Run(doc, "This text is read‑only."));

        // Insert the content control into the paragraph.
        paragraph.AppendChild(contentControl);

        // Ensure the output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "locked-content-control.docx");
        doc.Save(outputPath);
    }
}
