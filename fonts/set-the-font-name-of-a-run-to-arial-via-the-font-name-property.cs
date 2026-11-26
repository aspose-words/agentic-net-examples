using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a Run with some sample text.
        Run run = new Run(doc, "Sample text");

        // Set the font name of the Run to Arial.
        run.Font.Name = "Arial";

        // Validate that the font name was set correctly.
        if (run.Font.Name != "Arial")
        {
            throw new InvalidOperationException("Font name was not set to Arial.");
        }

        // Append the Run to the document.
        builder.Writeln(); // Ensure we are in a new paragraph.
        builder.CurrentParagraph.AppendChild(run);

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
