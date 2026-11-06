using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a paragraph and a run with some text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Hello, Aspose.Words!");

        // Set the font size of the run (in points).
        run.Font.Size = 16;

        // Add the run to the paragraph and the paragraph to the document.
        paragraph.AppendChild(run);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Retrieve the current font size of the run.
        double currentFontSize = run.Font.Size;

        // Output the font size for debugging purposes.
        Console.WriteLine($"Current Run font size: {currentFontSize}");

        // Save the document to ensure an output file is created.
        doc.Save("Output.docx");
    }
}
