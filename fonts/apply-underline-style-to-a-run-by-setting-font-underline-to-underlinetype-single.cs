using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add an initial paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph precedes the underlined run.");

        // Create a Run with the desired text.
        Run run = new Run(doc, "Underlined Run");

        // Apply single underline style to the run's font.
        run.Font.Underline = Aspose.Words.Underline.Single;

        // Validate that the underline was set correctly.
        if (run.Font.Underline != Aspose.Words.Underline.Single)
            throw new InvalidOperationException("Failed to set underline style.");

        // Append the run to the document's first paragraph.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "UnderlineRun.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
