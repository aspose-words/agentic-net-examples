using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline plain‑text content control at the current builder position.
        StructuredDocumentTag contentControl = builder.InsertStructuredDocumentTag(SdtType.PlainText);

        // Set the friendly title and tag for later identification.
        contentControl.Title = "CustomerName";
        contentControl.Tag = "CustomerNameTag";

        // Add placeholder text inside the content control.
        builder.Write("Enter name here");

        // Save the resulting document.
        doc.Save("ContentControlTitleTag.docx");
    }
}
