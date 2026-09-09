using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to host the run.
        if (doc.FirstSection?.Body?.FirstParagraph == null)
        {
            doc.FirstSection.Body.AppendChild(new Paragraph(doc));
        }

        // Create a run with sample text.
        Run run = new Run(doc, "Hello World!");

        // Set the font name of the run to Arial.
        Aspose.Words.Font font = run.Font;
        font.Name = "Arial";

        // Validate that the font name was set correctly.
        if (font.Name != "Arial")
            throw new InvalidOperationException("Failed to set the font name to Arial.");

        // Append the run to the first paragraph.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RunWithArial.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
