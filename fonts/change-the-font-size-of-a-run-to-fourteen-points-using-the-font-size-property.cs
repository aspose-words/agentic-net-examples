using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a run with some text.
        Run run = new Run(doc, "Sample text with 14pt font size.");

        // Change the font size of the run to 14 points.
        Aspose.Words.Font font = run.Font;
        font.Size = 14;

        // Validate that the font size was set correctly.
        if (Math.Abs(font.Size - 14) > 0.001)
            throw new InvalidOperationException("Font size was not set to 14 points.");

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSizeExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
