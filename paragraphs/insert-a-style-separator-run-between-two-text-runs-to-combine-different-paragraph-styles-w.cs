using System;
using System.IO;
using Aspose.Words;

public class StyleSeparatorExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First part of the line uses a built‑in heading style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator so the next text can have a different paragraph style
        // without breaking the line.
        builder.InsertStyleSeparator();

        // Define a custom paragraph style.
        Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 8;
        customStyle.Font.Bold = false;

        // Apply the custom style to the following text.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyleSeparator.docx");
        doc.Save(outputPath);
    }
}
