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

        // Add several paragraphs using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Set font color for each paragraph based on its index.
        for (int i = 0; i < doc.FirstSection.Body.Paragraphs.Count; i++)
        {
            Paragraph para = doc.FirstSection.Body.Paragraphs[i];

            // Even index -> Red, odd index -> Blue.
            Aspose.Drawing.Color aspColor = (i % 2 == 0) ? Aspose.Drawing.Color.Red : Aspose.Drawing.Color.Blue;

            // Convert Aspose.Drawing.Color to System.Drawing.Color for Font.Color.
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

            // Apply the color to every run in the paragraph.
            foreach (Run run in para.Runs)
            {
                run.Font.Color = sysColor;

                // Validate that the color was set correctly.
                if (run.Font.Color.ToArgb() != sysColor.ToArgb())
                {
                    throw new InvalidOperationException($"Failed to set color for paragraph {i + 1}");
                }
            }
        }

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicFontColors.docx");

        // Save the document.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
