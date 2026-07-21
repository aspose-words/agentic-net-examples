using System;
using Aspose.Words;
using Aspose.Drawing; // Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs to the document.
        const int paragraphCount = 9;
        for (int i = 0; i < paragraphCount; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Loop through each paragraph and set its font color based on the paragraph index.
        for (int i = 0; i < doc.FirstSection.Body.Paragraphs.Count; i++)
        {
            Paragraph para = doc.FirstSection.Body.Paragraphs[i];

            // Ensure the paragraph contains at least one Run before accessing it.
            if (para.Runs.Count == 0)
                continue; // Skip empty paragraphs.

            // Each paragraph created by DocumentBuilder contains at least one Run.
            var run = para.Runs[0];

            // Choose a color using Aspose.Drawing.Color.
            Aspose.Drawing.Color aspColor;
            switch (i % 3)
            {
                case 0:
                    aspColor = Aspose.Drawing.Color.Red;
                    break;
                case 1:
                    aspColor = Aspose.Drawing.Color.Green;
                    break;
                default:
                    aspColor = Aspose.Drawing.Color.Blue;
                    break;
            }

            // Convert Aspose.Drawing.Color to System.Drawing.Color as required by Font.Color.
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

            // Apply the color to the run's font.
            run.Font.Color = sysColor;

            // Validation: ensure the color was set correctly.
            if (run.Font.Color.ToArgb() != sysColor.ToArgb())
                throw new InvalidOperationException("Font color assignment failed.");
        }

        // Save the document to the local file system.
        string outputPath = "DynamicFontColors.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!System.IO.File.Exists(outputPath))
            throw new InvalidOperationException("Output file was not created.");
    }
}
