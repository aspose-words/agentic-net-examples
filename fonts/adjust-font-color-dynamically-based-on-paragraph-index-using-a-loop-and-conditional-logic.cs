using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs to the document.
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Loop through each paragraph and set its font color based on the paragraph index.
        for (int i = 0; i < doc.FirstSection.Body.Paragraphs.Count; i++)
        {
            Paragraph paragraph = doc.FirstSection.Body.Paragraphs[i];

            // Choose a color: even index -> Red, odd index -> Blue.
            Aspose.Drawing.Color aspColor = (i % 2 == 0) ? Aspose.Drawing.Color.Red : Aspose.Drawing.Color.Blue;

            // Convert Aspose.Drawing.Color to System.Drawing.Color as required by Font.Color.
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

            // Apply the color to every run within the paragraph.
            foreach (Run run in paragraph.Runs)
            {
                run.Font.Color = sysColor;

                // Validate that the color was set correctly.
                if (run.Font.Color.ToArgb() != sysColor.ToArgb())
                {
                    throw new InvalidOperationException("Failed to assign font color.");
                }
            }
        }

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicFontColors.docx");
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }
    }
}
