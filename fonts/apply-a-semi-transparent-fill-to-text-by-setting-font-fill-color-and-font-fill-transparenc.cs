using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a line of text.
        builder.Writeln("Hello, semi‑transparent world!");

        // Retrieve the first run (the text we just added).
        Run run = (Run)doc.FirstSection.Body.FirstParagraph.Runs[0];

        // Access the font of the run.
        Aspose.Words.Font font = run.Font;

        // Define a blue color using Aspose.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.FromArgb(255, 0, 0, 255); // Opaque blue.

        // Convert Aspose.Drawing.Color to System.Drawing.Color as required by the API.
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

        // Apply the fill color and set 50 % transparency.
        font.Fill.Color = sysColor;
        font.Fill.Transparency = 0.5; // 0 = opaque, 1 = fully transparent.

        // Validate that the properties were set correctly.
        bool colorMatches = font.Fill.Color.ToArgb() == sysColor.ToArgb();
        bool transparencyMatches = Math.Abs(font.Fill.Transparency - 0.5) < 0.0001;

        // Save the document to disk.
        string outputPath = "SemiTransparentText.docx";
        doc.Save(outputPath);

        // Ensure the file was created and properties are as expected.
        if (File.Exists(outputPath) && colorMatches && transparencyMatches)
        {
            // Execution completed successfully.
        }
    }
}
