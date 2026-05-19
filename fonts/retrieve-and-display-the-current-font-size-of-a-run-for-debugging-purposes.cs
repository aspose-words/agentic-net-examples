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
        Run run = new Run(doc, "Sample text");

        // Set a known font size for the run.
        Aspose.Words.Font font = run.Font;
        font.Size = 24; // points

        // Append the run to the document's first paragraph.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Retrieve the current font size from the run.
        double currentSize = run.Font.Size;

        // Display the font size for debugging.
        Console.WriteLine($"Current font size of the run: {currentSize} pt");

        // Save the document to ensure an output file is created.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RunFontSize.docx");
        doc.Save(outputPath);
    }
}
