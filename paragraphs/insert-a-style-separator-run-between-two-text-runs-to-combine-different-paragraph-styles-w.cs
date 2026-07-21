using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a built‑in style to the first part of the paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator so that the next text can have a different paragraph style
        // while staying on the same line.
        builder.InsertStyleSeparator();

        // Create a custom paragraph style.
        Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
        customStyle.Font.Bold = false;
        customStyle.Font.Size = 8;
        customStyle.Font.Name = "Arial";

        // Apply the custom style to the second part of the paragraph.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style.");

        // Save the document to a file.
        doc.Save("StyleSeparatorExample.docx");
    }
}
