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
        Run run = new Run(doc, "Bold, Italic, Underlined");

        // Apply bold, italic and underline formatting using the Font properties.
        run.Font.Bold = true;
        run.Font.Italic = true;
        run.Font.Underline = Aspose.Words.Underline.Single;

        // Validate that the properties were set correctly.
        if (!run.Font.Bold || !run.Font.Italic || run.Font.Underline != Aspose.Words.Underline.Single)
            throw new InvalidOperationException("Font formatting was not applied as expected.");

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormattedRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
