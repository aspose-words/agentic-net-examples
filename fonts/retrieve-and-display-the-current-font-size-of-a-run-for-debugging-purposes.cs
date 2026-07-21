using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a Run with sample text.
        Run run = new Run(doc, "Sample text for font size debugging.");

        // Set a known font size (in points) using the Run's Font property.
        Aspose.Words.Font font = run.Font;
        font.Size = 24;

        // Append the Run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Retrieve the current font size from the Run.
        double currentSize = run.Font.Size;

        // Output the font size to the console for debugging.
        Console.WriteLine($"Current font size of the run: {currentSize} points");

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSizeDebug.docx");
        doc.Save(outputPath);
    }
}
