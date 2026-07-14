using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Reusable method that applies a font name and size to a Run.
    public static void ApplyFont(Run run, string fontName, double fontSize)
    {
        // Set the font properties.
        run.Font.Name = fontName;
        run.Font.Size = fontSize;

        // Validate that the properties were applied correctly.
        if (run.Font.Name != fontName || Math.Abs(run.Font.Size - fontSize) > 0.001)
        {
            throw new InvalidOperationException("Failed to apply font settings to the run.");
        }
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph.
        Paragraph para = doc.FirstSection.Body.FirstParagraph ?? 
                         (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));

        // Create a run with sample text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Apply the desired font using the reusable method.
        ApplyFont(run, "Courier New", 24);

        // Add the run to the paragraph.
        para.AppendChild(run);

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormattedRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
