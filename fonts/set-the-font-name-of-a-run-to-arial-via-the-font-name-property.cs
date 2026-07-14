using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RunFontArial.docx");

        // Create a new blank document.
        Document doc = new Document();

        // Create a run with some sample text.
        Run run = new Run(doc, "This text uses Arial font.");

        // Set the font name of the run to Arial.
        run.Font.Name = "Arial";

        // Validate that the font name was set correctly.
        if (run.Font.Name != "Arial")
            throw new InvalidOperationException("Failed to set the run font to Arial.");

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
