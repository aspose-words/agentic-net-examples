using System;
using System.IO;
using System.Diagnostics;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a run with some text.
        Run run = new Run(doc, "Bold and Italic text");

        // Access the run's font and set Bold and Italic to true.
        Aspose.Words.Font font = run.Font;
        font.Bold = true;
        font.Italic = true;

        // Validate that the properties were set correctly.
        Debug.Assert(font.Bold, "Font.Bold should be true.");
        Debug.Assert(font.Italic, "Font.Italic should be true.");

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BoldItalicRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        Debug.Assert(File.Exists(outputPath), "Output file was not created.");
    }
}
