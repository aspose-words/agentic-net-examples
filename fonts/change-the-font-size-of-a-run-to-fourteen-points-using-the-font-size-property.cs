using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a Run with sample text.
        Run run = new Run(doc, "Sample text for font size change.");

        // Change the font size of the Run to 14 points.
        run.Font.Size = 14.0;

        // Insert the Run into the document.
        builder.InsertNode(run);

        // Validate that the font size was set correctly.
        if (run.Font.Size != 14.0)
        {
            throw new InvalidOperationException("Font size was not set to 14 points.");
        }

        // Save the document to a file.
        string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
