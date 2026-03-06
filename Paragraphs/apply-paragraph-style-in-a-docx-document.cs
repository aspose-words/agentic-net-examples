using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ApplyParagraphStyle
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Create a custom paragraph style named "MyCustomStyle".
        // ------------------------------------------------------------
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Calibri";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = System.Drawing.Color.DarkBlue;
        customStyle.ParagraphFormat.SpaceAfter = 12; // Add spacing after each paragraph.

        // ------------------------------------------------------------
        // 2. Apply the custom style to a paragraph using the builder.
        // ------------------------------------------------------------
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses the custom style \"MyCustomStyle\".");

        // ------------------------------------------------------------
        // 3. Apply an existing built‑in style by name to another paragraph.
        // ------------------------------------------------------------
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("This paragraph uses the built‑in \"Heading 1\" style.");

        // ------------------------------------------------------------
        // 4. Demonstrate style separator: two different styles on the same line.
        // ------------------------------------------------------------
        // First part of the line – use Heading 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Write("Part 1 (Heading 2) ");
        // Insert a style separator so the next text can have a different style.
        builder.InsertStyleSeparator();

        // Second part of the line – use Normal style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Part 2 (Normal).");

        // ------------------------------------------------------------
        // Save the document to a DOCX file.
        // ------------------------------------------------------------
        string outputPath = "AppliedParagraphStyles.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
