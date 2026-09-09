using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an inline plain‑text content control (StructuredDocumentTag).
        StructuredDocumentTag contentControl = new StructuredDocumentTag(
            doc,
            SdtType.PlainText,
            MarkupLevel.Inline);

        // Set the friendly title and the tag for later identification.
        contentControl.Title = "CustomerName";
        contentControl.Tag = "customer-name";

        // Add placeholder text inside the content control.
        contentControl.RemoveAllChildren();
        contentControl.AppendChild(new Run(doc, "Enter name here"));

        // Insert the content control into the document.
        builder.InsertNode(contentControl);

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ContentControlTitleTag.docx");
        doc.Save(outputPath);
    }
}
