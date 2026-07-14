using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Define a folder for the output document.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document and a DocumentBuilder to edit it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1 - level 0");

        // Increase the list level (indent) to create a sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 2 - level 1");

        // Increase the list level again for a deeper sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 3 - level 2");

        // Decrease the list level (outdent) back to the previous level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 4 - back to level 1");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output folder.
        string outputPath = Path.Combine(artifactsDir, "IncreaseIndent.docx");
        doc.Save(outputPath);
    }
}
