using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (exists by default).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Obtain a Font object from DocumentBuilder for configuration.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Aspose.Words.Font sharedFont = builder.Font;
        sharedFont.Size = 24; // Set font size.

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());
        sharedFont.Color = sysColor; // Assign the color to the font.

        // Helper method to create a Run with the shared font settings.
        Run CreateRun(string text)
        {
            Run run = new Run(doc, text);
            run.Font.Size = sharedFont.Size;
            run.Font.Color = sharedFont.Color;
            return run;
        }

        // Add multiple runs that share the same font formatting.
        paragraph.AppendChild(CreateRun("Hello "));
        paragraph.AppendChild(CreateRun("world"));
        paragraph.AppendChild(CreateRun("!"));

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "FontExample.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
