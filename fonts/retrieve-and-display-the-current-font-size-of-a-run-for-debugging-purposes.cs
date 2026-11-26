using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a new run with some text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Optionally set a specific font size (e.g., 14 points).
        // If not set, the default size will be used.
        run.Font.Size = 14.0;

        // Add the run to the first paragraph of the document.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(run);

        // Retrieve the current font size of the run.
        double currentFontSize = run.Font.Size;

        // Output the font size for debugging purposes.
        Console.WriteLine($"Run font size: {currentFontSize} pt");

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
