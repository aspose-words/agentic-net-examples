using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a single run of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text with fill color and transparency.");

        // Retrieve the first run that we just added.
        Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

        // Access the Fill formatting of the run's font.
        Fill fill = run.Font.Fill;

        // Ensure the fill type is solid before setting color and transparency.
        fill.Solid();

        // Set the foreground color to red.
        // Fill.Color expects System.Drawing.Color, so convert from Aspose.Drawing.Color.
        fill.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Red.ToArgb());

        // Set transparency to 30% (0.3). 0.0 = opaque, 1.0 = fully transparent.
        fill.Transparency = 0.3;

        // Validate that the properties were applied correctly.
        var expectedColor = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Red.ToArgb());
        if (fill.Color.ToArgb() != expectedColor.ToArgb() ||
            Math.Abs(fill.Transparency - 0.3) > 0.0001)
        {
            throw new InvalidOperationException("Fill properties were not set as expected.");
        }

        // Define the output file path.
        string outputPath = "FilledText.docx";

        // Save the document to disk.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
