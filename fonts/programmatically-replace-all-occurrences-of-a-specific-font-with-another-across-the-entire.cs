using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with mixed fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "OldFont";
        builder.Writeln("This paragraph uses OldFont.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses Times New Roman.");

        builder.Font.Name = "OldFont";
        builder.Writeln("Another line with OldFont.");

        // Save the initial document (optional, demonstrates load/save lifecycle).
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Define fonts to replace.
        string oldFontName = "OldFont";
        string newFontName = "NewFont";

        // Replace all occurrences of the old font with the new font.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (string.Equals(run.Font.Name, oldFontName, StringComparison.OrdinalIgnoreCase))
            {
                run.Font.Name = newFontName;
            }
        }

        // Validate that no run still uses the old font.
        bool replacementSuccessful = true;
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (string.Equals(run.Font.Name, oldFontName, StringComparison.OrdinalIgnoreCase))
            {
                replacementSuccessful = false;
                break;
            }
        }

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Ensure the output file exists and report the result.
        if (File.Exists(outputPath) && replacementSuccessful)
        {
            Console.WriteLine("Font replacement completed successfully.");
        }
        else
        {
            Console.WriteLine("Font replacement failed.");
        }
    }
}
