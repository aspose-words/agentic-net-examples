using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Define output path for the generated document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicFontColors.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs with placeholder text.
        int paragraphCount = 5;
        for (int i = 0; i < paragraphCount; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Loop through each paragraph and set the font color based on its index.
        for (int i = 0; i < doc.FirstSection.Body.Paragraphs.Count; i++)
        {
            var paragraph = doc.FirstSection.Body.Paragraphs[i];

            // Ensure the paragraph contains at least one run to apply formatting.
            if (paragraph.Runs.Count > 0)
            {
                var run = paragraph.Runs[0];

                // Choose a color: even index -> Red, odd index -> Blue.
                Aspose.Drawing.Color aspColor = (i % 2 == 0) ? Aspose.Drawing.Color.Red : Aspose.Drawing.Color.Blue;

                // Convert Aspose.Drawing.Color to System.Drawing.Color as required by the Font.Color property.
                System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

                // Apply the color to the run's font.
                run.Font.Color = sysColor;
            }
        }

        // Save the document to the specified file.
        doc.Save(outputPath);
    }
}
