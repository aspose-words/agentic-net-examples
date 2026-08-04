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

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Obtain the Aspose.Words.Font object from the builder.
        Aspose.Words.Font sharedFont = builder.Font;

        // Set common font properties.
        sharedFont.Size = 18; // Font size in points.

        // Convert Aspose.Drawing.Color to System.Drawing.Color for the Font.Color property.
        sharedFont.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Red.ToArgb());

        // Validate that the properties were set correctly.
        if (sharedFont.Size != 18 ||
            sharedFont.Color.ToArgb() != System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Red.ToArgb()).ToArgb())
        {
            Console.WriteLine("Font properties were not set as expected.");
            return;
        }

        // Insert multiple runs; they inherit the shared font formatting.
        builder.Writeln("First run with shared font.");
        builder.Writeln("Second run with shared font.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormattedRuns.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
