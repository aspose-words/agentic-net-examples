using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a new paragraph to the document.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a Run with sample text.
        Run run = new Run(doc, "Hello, Aspose.Words!");
        paragraph.AppendChild(run);

        // Set the font name of the Run to Arial.
        run.Font.Name = "Arial";

        // Validate that the font name was set correctly.
        if (run.Font.Name != "Arial")
        {
            throw new InvalidOperationException("Font name was not set to Arial.");
        }

        // Define the output file path.
        string outputPath = "Output.docx";

        // Save the document to disk.
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
