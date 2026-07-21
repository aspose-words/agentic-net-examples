using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Aspose.Words.Document doc = new Aspose.Words.Document();

        // The document already contains one section, body, and paragraph.
        Aspose.Words.Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some text.
        Aspose.Words.Run run = new Aspose.Words.Run(doc, "Bold, Italic, Underlined text");

        // Apply bold, italic, and underline formatting to the run.
        Aspose.Words.Font runFont = run.Font;
        runFont.Bold = true;
        runFont.Italic = true;
        runFont.Underline = Aspose.Words.Underline.Single;

        // Validate that the formatting was applied correctly.
        if (!runFont.Bold || !runFont.Italic || runFont.Underline != Aspose.Words.Underline.Single)
        {
            throw new InvalidOperationException("Font formatting was not applied as expected.");
        }

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document to a file.
        const string outputPath = "FormattedRun.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved.", outputPath);
        }
    }
}
