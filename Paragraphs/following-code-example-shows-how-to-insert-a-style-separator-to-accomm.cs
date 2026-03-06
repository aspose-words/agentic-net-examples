using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the first part of the line to a built‑in heading style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator so the next text can have a different paragraph style
        // while staying on the same visual line.
        builder.InsertStyleSeparator();

        // Define a custom paragraph style.
        Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
        customStyle.Font.Bold = false;
        customStyle.Font.Size = 8;
        customStyle.Font.Name = "Arial";

        // Apply the custom style and write the second part of the line.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style. ");

        // Save the resulting document.
        doc.Save("StyleSeparatorExample.docx");
    }
}
