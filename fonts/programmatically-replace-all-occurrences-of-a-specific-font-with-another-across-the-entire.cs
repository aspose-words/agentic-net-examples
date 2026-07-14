using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document with text using two different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "Arial";
        builder.Writeln("This line uses Arial font.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line uses Times New Roman font.");

        // Replace all occurrences of the source font ("Arial") with the target font ("Courier New").
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (string.Equals(run.Font.Name, "Arial", StringComparison.OrdinalIgnoreCase))
                run.Font.Name = "Courier New";
        }

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "FontReplaced.docx");
        doc.Save(outputPath);
    }
}
