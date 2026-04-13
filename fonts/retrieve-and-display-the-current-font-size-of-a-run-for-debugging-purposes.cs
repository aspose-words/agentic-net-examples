using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Initialize a new blank document.
        Document doc = new Document();

        // Get the first paragraph (created by default).
        var paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with sample text and assign a font size.
        Run run = new Run(doc, "Sample text");
        run.Font.Size = 24; // Font size in points.

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document to satisfy validation requirements.
        string docPath = Path.Combine(artifactsDir, "RunFontSize.docx");
        doc.Save(docPath);

        // Retrieve the current font size of the run.
        double currentSize = run.Font.Size;

        // Display the font size for debugging purposes.
        Console.WriteLine($"Run font size: {currentSize} points");
    }
}
