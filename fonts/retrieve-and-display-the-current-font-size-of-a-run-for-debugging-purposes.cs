using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (created by default) to hold the run.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some text.
        Run run = new Run(doc, "Sample text");

        // Set a known font size (in points) for the run.
        run.Font.Size = 24;

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document so we can verify that the file is created.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "RunFontSizeDemo.docx");
        doc.Save(outputPath);

        // Retrieve the current font size from the run.
        double currentSize = run.Font.Size;

        // Output the font size for debugging purposes.
        Console.WriteLine($"Run font size: {currentSize} points");
    }
}
