using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Reusable method that applies a font name and size to a Run.
    public static void ApplyFont(Run run, string fontName, double fontSize)
    {
        // Use Aspose.Words.Font for document text formatting.
        Aspose.Words.Font font = run.Font;
        font.Name = fontName;
        font.Size = fontSize;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert a Run.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create the Run node.
        Run run = new Run(doc, "This is a sample text.");

        // Add the Run to the current paragraph and keep a reference.
        builder.CurrentParagraph.Runs.Add(run);

        // Apply the desired font using the reusable method.
        ApplyFont(run, "Arial", 16);

        // Validate that the font properties were set correctly.
        if (run.Font.Name != "Arial" || Math.Abs(run.Font.Size - 16) > 0.01)
        {
            throw new InvalidOperationException("Font properties were not applied correctly.");
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }

        // Indicate successful completion (no interactive prompts).
        Console.WriteLine("Document created and saved successfully.");
    }
}
