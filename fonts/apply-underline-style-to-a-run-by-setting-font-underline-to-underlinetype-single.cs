using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to create an empty paragraph where the run will be placed.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(); // adds a new empty paragraph.

        // Create a Run with the desired text.
        Run run = new Run(doc, "Underlined text.");

        // Apply a single underline style to the run's font.
        run.Font.Underline = Aspose.Words.Underline.Single;

        // Validate that the underline was set correctly.
        if (run.Font.Underline != Aspose.Words.Underline.Single)
            throw new InvalidOperationException("Failed to set underline style on the run.");

        // Append the run to the first paragraph of the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(run);

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "UnderlineRun.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);

        // Indicate successful completion.
        Console.WriteLine("Document created with underlined run at: " + outputPath);
    }
}
