using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some text with different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text using the font that we want to replace.
        builder.Font.Name = "OldFont";
        builder.Writeln("This text uses the OldFont.");

        // Text using another font (should remain unchanged).
        builder.Font.Name = "AnotherFont";
        builder.Writeln("This text uses AnotherFont.");

        // Replace all occurrences of the old font with the new font.
        const string fontToReplace = "OldFont";
        const string replacementFont = "NewFont";

        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (string.Equals(run.Font.Name, fontToReplace, StringComparison.OrdinalIgnoreCase))
            {
                run.Font.Name = replacementFont;
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "FontReplaced.docx");
        doc.Save(outputPath);
    }
}
