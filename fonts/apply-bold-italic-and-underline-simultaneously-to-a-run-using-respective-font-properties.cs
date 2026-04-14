using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        if (paragraph == null)
        {
            paragraph = new Paragraph(doc);
            doc.FirstSection.Body.AppendChild(paragraph);
        }

        // Create a run with sample text.
        Run run = new Run(doc, "Bold, Italic and Underlined text");

        // Access the font of the run using the fully qualified type.
        Aspose.Words.Font runFont = run.Font;

        // Apply bold, italic and underline formatting.
        runFont.Bold = true;
        runFont.Italic = true;
        runFont.Underline = Underline.Single;

        // Validate that the formatting was applied correctly.
        if (!runFont.Bold || !runFont.Italic || runFont.Underline != Underline.Single)
            throw new InvalidOperationException("Font formatting was not applied as expected.");

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BoldItalicUnderline.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
