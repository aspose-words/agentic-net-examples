using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define the font names to replace.
        const string oldFontName = "Arial";
        const string newFontName = "Times New Roman";

        // Create a new document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = oldFontName;
        builder.Writeln("This paragraph uses the old font.");

        builder.Font.Name = "Courier New";
        builder.Writeln("This paragraph uses a different font.");

        // Replace all occurrences of the old font with the new font.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (string.Equals(run.Font.Name, oldFontName, StringComparison.OrdinalIgnoreCase))
                run.Font.Name = newFontName;
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
