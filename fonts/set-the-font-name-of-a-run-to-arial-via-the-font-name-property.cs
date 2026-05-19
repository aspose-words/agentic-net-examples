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
        Run run = new Run(doc, "Hello World!");

        // Set the run's font name to Arial.
        Aspose.Words.Font font = run.Font;
        font.Name = "Arial";

        // Validate that the font name was set correctly.
        if (font.Name != "Arial")
            throw new InvalidOperationException("Font name was not set to Arial.");

        // Append the run to the first paragraph of the document.
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
