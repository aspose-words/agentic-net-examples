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

        // Apply a built‑in paragraph style to the first part of the line.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator so that the next text can have a different paragraph style
        // while staying on the same visual line.
        builder.InsertStyleSeparator();

        // Create a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyParaStyle");
        customStyle.Font.Bold = false;
        customStyle.Font.Size = 8;
        customStyle.Font.Name = "Arial";

        // Apply the custom style to the second part of the line.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style. ");

        // Save the document to the local file system.
        doc.Save("StyleSeparatorExample.docx");
    }
}
