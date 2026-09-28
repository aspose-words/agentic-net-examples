using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a new paragraph to the document body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a Run with sample text.
        Run run = new Run(doc, "Sample text for font size debugging.");

        // Set the font size of the Run (in points).
        run.Font.Size = 14;

        // Add the Run to the paragraph.
        paragraph.AppendChild(run);

        // Retrieve the current font size of the Run.
        double currentFontSize = run.Font.Size;

        // Output the font size to the console.
        Console.WriteLine($"Current Run font size: {currentFontSize} points");

        // Save the document to verify that the output file exists.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
