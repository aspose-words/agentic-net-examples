using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Ensure the document has a paragraph to host the run.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        if (paragraph == null)
        {
            paragraph = new Paragraph(doc);
            doc.FirstSection.Body.AppendChild(paragraph);
        }

        // Create a run with sample text.
        Run run = new Run(doc, "Sample text with custom font.");

        // Apply the desired font name and size using the reusable method.
        ApplyFont(run, "Courier New", 24);

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Prepare output directory and file path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CustomFontRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
            Console.WriteLine("Document saved successfully: " + outputPath);
        else
            Console.WriteLine("Failed to save the document.");
    }

    /// <summary>
    /// Applies a specific font name and size to the provided Run.
    /// </summary>
    /// <param name="run">The Run whose font will be modified.</param>
    /// <param name="fontName">The name of the font to apply.</param>
    /// <param name="fontSize">The size of the font in points.</param>
    public static void ApplyFont(Run run, string fontName, double fontSize)
    {
        // Set font properties using Aspose.Words.Font.
        run.Font.Name = fontName;
        run.Font.Size = fontSize;

        // Validate that the properties were set correctly.
        if (run.Font.Name != fontName || Math.Abs(run.Font.Size - fontSize) > 0.001)
            throw new InvalidOperationException("Failed to apply font settings to the run.");
    }
}
