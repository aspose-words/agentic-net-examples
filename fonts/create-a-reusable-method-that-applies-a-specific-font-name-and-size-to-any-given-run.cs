using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Reusable method that applies a specific font name and size to a Run.
    public static void ApplyFont(Run run, string fontName, double fontSize)
    {
        // Apply font properties.
        run.Font.Name = fontName;
        run.Font.Size = fontSize;

        // Validate that the properties were set correctly.
        if (run.Font.Name != fontName || Math.Abs(run.Font.Size - fontSize) > 0.001)
        {
            throw new InvalidOperationException("Failed to apply font properties to the Run.");
        }
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a single run.
        builder.Writeln("This is a sample text.");

        // After Writeln the builder moves to a new empty paragraph.
        // Retrieve the paragraph that contains the text (the previous sibling).
        Paragraph textParagraph = builder.CurrentParagraph.PreviousSibling as Paragraph;
        if (textParagraph == null || textParagraph.Runs.Count == 0)
        {
            throw new InvalidOperationException("The expected paragraph or run was not found.");
        }

        // Retrieve the first run in that paragraph.
        Run firstRun = textParagraph.Runs[0];

        // Apply the desired font to the run.
        ApplyFont(firstRun, "Arial", 16);

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        // Indicate success (no interactive prompts).
        Console.WriteLine("Document created successfully at: " + Path.GetFullPath(outputPath));
    }
}
