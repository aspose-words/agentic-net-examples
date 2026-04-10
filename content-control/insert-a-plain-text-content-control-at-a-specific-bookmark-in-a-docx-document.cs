using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a bookmark named "TargetBookmark".
        builder.Writeln("Paragraph before the bookmark.");
        builder.StartBookmark("TargetBookmark");
        // The bookmark can be empty; we will insert the content control at this position.
        builder.EndBookmark("TargetBookmark");
        builder.Writeln("Paragraph after the bookmark.");

        // Move the cursor to the start of the bookmark.
        builder.MoveToBookmark("TargetBookmark");

        // Insert a plain‑text structured document tag (content control) at the bookmark.
        StructuredDocumentTag plainTextSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);

        // Optional: set metadata for the content control.
        plainTextSdt.Title = "SamplePlainTextControl";
        plainTextSdt.Tag = "PlainTextTag";

        // Add some default text inside the content control.
        plainTextSdt.AppendChild(new Run(doc, "Editable content inside the control."));

        // Save the resulting document to the working directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputWithContentControl.docx");
        doc.Save(outputPath);
    }
}
