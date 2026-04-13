using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledParagraphs.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Create a custom paragraph style with specific font settings.
        // -----------------------------------------------------------------
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = System.Drawing.Color.DarkBlue; // Fully qualified System.Drawing.Color
        customStyle.Font.Bold = true;

        // -----------------------------------------------------------------
        // Add a range of paragraphs and apply the custom style to each.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 5; i++)
        {
            // Apply the style to the current paragraph.
            builder.ParagraphFormat.Style = customStyle;
            builder.Writeln($"Paragraph {i} using the custom style.");
        }

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // Validate that the style's font properties are as expected.
        // -----------------------------------------------------------------
        bool isValid =
            customStyle.Font.Name == "Arial" &&
            Math.Abs(customStyle.Font.Size - 14) < 0.01 &&
            customStyle.Font.Bold &&
            customStyle.Font.Color.ToArgb() == System.Drawing.Color.DarkBlue.ToArgb();

        if (!isValid)
            throw new InvalidOperationException("Style font properties validation failed.");

        // -----------------------------------------------------------------
        // Ensure the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
