using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a built‑in style to the first part of the text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator so the next text can have a different paragraph style.
        builder.InsertStyleSeparator();

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyParaStyle");
        customStyle.Font.Size = 8;
        customStyle.Font.Name = "Arial";
        customStyle.Font.Bold = false;

        // Apply the custom style and write the second part of the text.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style.");

        // Save the resulting document.
        doc.Save("StyleSeparator.docx");
    }
}
