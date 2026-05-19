using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a run with some text.
        Run run = new Run(doc, "Underlined text");

        // Apply a single underline to the run's font.
        run.Font.Underline = Aspose.Words.Underline.Single;

        // Append the run to the current paragraph.
        builder.CurrentParagraph.AppendChild(run);
        builder.Writeln(); // Add a line break.

        // Validate that the underline was set correctly.
        if (run.Font.Underline != Aspose.Words.Underline.Single)
            throw new InvalidOperationException("Underline style was not applied.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UnderlineRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Document was not saved.", outputPath);

        // Indicate success.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
