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

        // Create a run with sample text.
        Run run = new Run(doc, "Bold and Italic text");

        // Apply bold and italic formatting to the run's font.
        run.Font.Bold = true;
        run.Font.Italic = true;

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Define the output file path.
        string outputPath = "BoldItalicRun.docx";

        // Save the document to disk.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
