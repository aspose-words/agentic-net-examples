using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // Provides Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a few Aspose.Drawing colors to choose from.
        Aspose.Drawing.Color redAspose = Aspose.Drawing.Color.Red;
        Aspose.Drawing.Color greenAspose = Aspose.Drawing.Color.Green;
        Aspose.Drawing.Color blueAspose = Aspose.Drawing.Color.Blue;

        // Add several paragraphs, changing the font color based on the paragraph index.
        for (int i = 0; i < 5; i++)
        {
            // Select a color using conditional logic.
            Aspose.Drawing.Color selectedAsposeColor;
            if (i % 3 == 0)
                selectedAsposeColor = redAspose;
            else if (i % 3 == 1)
                selectedAsposeColor = greenAspose;
            else
                selectedAsposeColor = blueAspose;

            // Convert Aspose.Drawing.Color to System.Drawing.Color as required by Font.Color.
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(selectedAsposeColor.ToArgb());

            // Apply the color to the current font.
            builder.Font.Color = sysColor;

            // Write a paragraph.
            builder.Writeln($"Paragraph {i + 1} with dynamically set color.");

            // Validate that the color was set correctly.
            if (builder.Font.Color.ToArgb() != sysColor.ToArgb())
                throw new InvalidOperationException("Font color was not applied as expected.");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicFontColors.docx");
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
