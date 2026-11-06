using System;
using System.IO;
using Aspose.Words;

public class FontApplier
{
    // Reusable method that applies a specific font name and size to a Run.
    public static void ApplyFont(Run run, string fontName, double fontSize)
    {
        // Apply font name and size.
        run.Font.Name = fontName;
        run.Font.Size = fontSize;

        // Validate that the properties were set correctly.
        if (run.Font.Name != fontName || Math.Abs(run.Font.Size - fontSize) > 0.01)
        {
            throw new InvalidOperationException("Failed to apply font properties to the Run.");
        }
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a Run.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample text that will have its font changed.");

        // Retrieve the first Run in the document.
        Run firstRun = (Run)doc.GetChild(NodeType.Run, 0, true);
        if (firstRun == null)
        {
            throw new InvalidOperationException("No Run found in the document.");
        }

        // Apply the desired font name and size.
        ApplyFont(firstRun, "Arial", 16);

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
