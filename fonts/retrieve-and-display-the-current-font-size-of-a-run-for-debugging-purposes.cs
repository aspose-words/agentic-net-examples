using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a Run with some text.
        Run run = new Run(doc, "Sample text for font size debugging.");

        // Set a known font size for the run.
        run.Font.Size = 24.0; // points

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Retrieve the current font size of the run.
        double currentFontSize = run.Font.Size;

        // Output the font size to the console.
        Console.WriteLine($"Current Run Font Size: {currentFontSize} points");

        // Save the document to verify that the run was added correctly.
        string outputPath = "RunFontSizeDebug.docx";
        doc.Save(outputPath);
    }
}
