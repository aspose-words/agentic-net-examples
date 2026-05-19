using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a run with the desired text.
        Run run = new Run(doc, "Bold and Italic text");

        // Apply bold and italic formatting to the run's font.
        Aspose.Words.Font font = run.Font;
        font.Bold = true;
        font.Italic = true;

        // Validate that the formatting was applied.
        if (!font.Bold || !font.Italic)
        {
            throw new InvalidOperationException("Failed to apply bold or italic formatting.");
        }

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BoldItalicRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved.", outputPath);
        }

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
