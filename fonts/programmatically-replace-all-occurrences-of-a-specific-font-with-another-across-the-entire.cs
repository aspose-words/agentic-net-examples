using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text using two different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "Arial";
        builder.Writeln("This line is formatted with Arial.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line is formatted with Times New Roman.");

        builder.Font.Name = "Arial";
        builder.Writeln("Another line in Arial.");

        // Define the font to replace and the replacement font.
        const string oldFontName = "Arial";
        const string newFontName = "Courier New";

        // Iterate through all Run nodes and replace the font name where it matches the old font.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true).Cast<Run>())
        {
            if (string.Equals(run.Font.Name, oldFontName, StringComparison.OrdinalIgnoreCase))
            {
                run.Font.Name = newFontName;
            }
        }

        // Optional validation: ensure that at least one run now uses the new font.
        bool replacementOccurred = doc.GetChildNodes(NodeType.Run, true)
            .Cast<Run>()
            .Any(r => string.Equals(r.Font.Name, newFontName, StringComparison.OrdinalIgnoreCase));

        // Save the modified document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ReplacedFont.docx");
        doc.Save(outputPath);
    }
}
