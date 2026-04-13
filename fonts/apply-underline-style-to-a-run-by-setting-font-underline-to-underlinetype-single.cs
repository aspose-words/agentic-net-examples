using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a run with sample text.
        Run run = new Run(doc, "Underlined text");

        // Apply a single underline style to the run's font.
        run.Font.Underline = Underline.Single;

        // Validate that the underline property was set correctly.
        if (run.Font.Underline != Underline.Single)
            throw new InvalidOperationException("Failed to set underline style.");

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "UnderlineRun.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Ensure that the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
