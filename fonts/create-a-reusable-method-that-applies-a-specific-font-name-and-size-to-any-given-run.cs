using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to hold the run.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with some sample text.
        Run run = new Run(doc, "Sample text for font formatting.");

        // Apply the desired font name and size using the reusable method.
        ApplyFont(run, "Courier New", 24);

        // Add the formatted run to the paragraph.
        paragraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormattedRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }

    /// <summary>
    /// Applies a specific font name and size to the provided Run.
    /// </summary>
    /// <param name="run">The Run to format.</param>
    /// <param name="fontName">The name of the font to apply.</param>
    /// <param name="fontSize">The size of the font in points.</param>
    public static void ApplyFont(Run run, string fontName, double fontSize)
    {
        if (run == null) throw new ArgumentNullException(nameof(run));

        // Access the Font object of the run.
        Aspose.Words.Font font = run.Font;

        // Set font properties.
        font.Name = fontName;
        font.Size = fontSize;

        // Validate that the properties were set correctly.
        if (font.Name != fontName || Math.Abs(font.Size - fontSize) > 0.001)
        {
            throw new InvalidOperationException("Font properties were not applied correctly.");
        }
    }
}
