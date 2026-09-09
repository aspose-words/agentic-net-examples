using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define two colors using Aspose.Drawing.Color.
        Aspose.Drawing.Color blueAspose = Aspose.Drawing.Color.Blue;
        Aspose.Drawing.Color redAspose = Aspose.Drawing.Color.Red;

        // Convert Aspose.Drawing.Color to System.Drawing.Color for the Font.Color property.
        System.Drawing.Color blue = System.Drawing.Color.FromArgb(blueAspose.ToArgb());
        System.Drawing.Color red = System.Drawing.Color.FromArgb(redAspose.ToArgb());

        // Add several paragraphs, changing the font color based on the paragraph index.
        for (int i = 0; i < 5; i++)
        {
            // Even index -> blue, odd index -> red.
            System.Drawing.Color currentColor = (i % 2 == 0) ? blue : red;

            // Apply the selected color to the builder's font.
            builder.Font.Color = currentColor;

            // Write the paragraph text.
            builder.Writeln($"Paragraph {i + 1} with {(i % 2 == 0 ? "blue" : "red")} text.");

            // Validate that the color was set correctly.
            if (builder.Font.Color.ToArgb() != currentColor.ToArgb())
                throw new InvalidOperationException("Font color assignment validation failed.");
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DynamicFontColors.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Document was not saved correctly.", outputPath);
    }
}
