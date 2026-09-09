using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a Run containing the desired text.
        Run run = new Run(doc, "Bold and Italic text");

        // Access the Run's font and set Bold and Italic to true.
        Aspose.Words.Font font = run.Font;
        font.Bold = true;
        font.Italic = true;

        // Validate that the properties were set correctly.
        if (!font.Bold || !font.Italic)
            throw new InvalidOperationException("Failed to apply bold or italic style to the run.");

        // Append the Run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BoldItalicRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
