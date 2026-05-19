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

        // Ensure the document has at least one paragraph to host the run.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some text.
        Run run = new Run(doc, "Bold, Italic, Underlined text");

        // Apply bold, italic, and underline formatting via the Run's Font object.
        Aspose.Words.Font font = run.Font;
        font.Bold = true;
        font.Italic = true;
        font.Underline = Underline.Single;

        // Validate that the properties were set correctly.
        if (!font.Bold || !font.Italic || font.Underline != Underline.Single)
            throw new InvalidOperationException("Failed to set font formatting.");

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormattedRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
