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
        Run run = new Run(doc, "Sample text");

        // Change the font size of the run to 14 points.
        run.Font.Size = 14;

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Validate that the font size was set correctly.
        if (run.Font.Size != 14)
            throw new InvalidOperationException("Font size was not set to 14 points.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RunFontSize.docx");

        // Save the document to the file system.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
