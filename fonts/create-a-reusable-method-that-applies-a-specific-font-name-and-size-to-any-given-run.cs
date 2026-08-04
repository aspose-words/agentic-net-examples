using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Reusable method that applies a specific font name and size to a Run.
    public static void ApplyFont(Run run, string fontName, double fontSize)
    {
        if (run == null) throw new ArgumentNullException(nameof(run));

        // Set font properties.
        run.Font.Name = fontName;
        run.Font.Size = fontSize;

        // Validate that the properties were applied correctly.
        if (run.Font.Name != fontName || Math.Abs(run.Font.Size - fontSize) > 0.001)
            throw new InvalidOperationException("Failed to set font properties on the run.");
    }

    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Ensure the document has a paragraph to hold the run.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with sample text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Apply the desired font using the reusable method.
        ApplyFont(run, "Courier New", 24);

        // Add the run to the paragraph.
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
}
