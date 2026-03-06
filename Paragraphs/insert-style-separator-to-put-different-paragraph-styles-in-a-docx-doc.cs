using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First part of the line uses a built‑in heading style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator so the next text can have a different style
        // but remain on the same visual line.
        builder.InsertStyleSeparator();

        // Define a custom paragraph style.
        Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
        customStyle.Font.Bold = false;
        customStyle.Font.Size = 8;
        customStyle.Font.Name = "Arial";

        // Apply the custom style to the second part of the line.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style.");

        // Save the document as DOCX.
        doc.Save("StyleSeparatorExample.docx");
    }
}
