using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a run with sample text.
        Run run = new Run(doc, "Bold, Italic, Underlined text");

        // Access the run's font and apply bold, italic, and underline formatting.
        Aspose.Words.Font font = run.Font;
        font.Bold = true;
        font.Italic = true;
        font.Underline = Underline.Single;

        // Validate that the formatting was applied correctly.
        if (!font.Bold || !font.Italic || font.Underline != Underline.Single)
            throw new InvalidOperationException("Font formatting was not applied as expected.");

        // Append the run to the document's first paragraph.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormattedRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
