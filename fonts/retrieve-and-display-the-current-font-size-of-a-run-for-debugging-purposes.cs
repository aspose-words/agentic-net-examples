using System;
using System.IO;
using Aspose.Words;

public class RetrieveRunFontSize
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a run with sample text.
        Run run = new Run(doc, "Sample text");

        // Set a specific font size for the run.
        run.Font.Size = 24;

        // Append the run to the first paragraph of the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(run);

        // Save the document (optional, but ensures output file exists).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "RunFontSize.docx");
        doc.Save(outputPath);

        // Retrieve the current font size of the run.
        double currentSize = run.Font.Size;

        // Display the font size for debugging purposes.
        Console.WriteLine($"Run font size: {currentSize} points");
    }
}
