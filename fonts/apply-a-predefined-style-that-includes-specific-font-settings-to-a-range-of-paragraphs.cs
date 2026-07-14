using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class ApplyStyleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Define a custom paragraph style with specific font settings.
        // -----------------------------------------------------------------
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Courier New";
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;
        // Font.Color expects System.Drawing.Color. Convert from Aspose.Drawing.Color.
        customStyle.Font.Color = System.Drawing.Color.FromArgb(Color.Blue.ToArgb());

        // Validate that the style's font properties were set correctly.
        if (customStyle.Font.Name != "Courier New")
            throw new InvalidOperationException("Font name was not set correctly on the style.");
        if (customStyle.Font.Size != 14)
            throw new InvalidOperationException("Font size was not set correctly on the style.");
        if (!customStyle.Font.Bold)
            throw new InvalidOperationException("Font bold flag was not set correctly on the style.");
        System.Drawing.Color expectedColor = System.Drawing.Color.FromArgb(Color.Blue.ToArgb());
        if (!customStyle.Font.Color.Equals(expectedColor))
            throw new InvalidOperationException("Font color was not set correctly on the style.");

        // -----------------------------------------------------------------
        // Insert paragraphs. Apply the custom style to a specific range.
        // -----------------------------------------------------------------
        builder.Writeln("Paragraph without custom style.");

        // Apply the custom style to the next three paragraphs.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("Paragraph with custom style 1.");
        builder.Writeln("Paragraph with custom style 2.");
        builder.Writeln("Paragraph with custom style 3.");

        // Reset to the default style for subsequent paragraphs.
        builder.ParagraphFormat.Style = doc.Styles["Normal"];
        builder.Writeln("Paragraph after custom style.");

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledParagraphs.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
