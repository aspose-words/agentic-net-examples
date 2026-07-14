using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable line numbering for the whole section.
        // Numbers will start at 1 and increase by 1 for each line.
        builder.PageSetup.LineStartingNumber = 1;
        builder.PageSetup.LineNumberCountBy = 1;
        builder.PageSetup.LineNumberRestartMode = LineNumberRestartMode.Continuous;

        // Add several paragraphs to the document.
        builder.Writeln("First paragraph of the document.");
        builder.Writeln("Second paragraph, containing more text.");
        builder.Writeln("Third paragraph, demonstrating line numbering.");
        builder.Writeln("Fourth paragraph, final example.");

        // Configure save options for plain‑text export.
        // Line numbers are automatically included because line numbering is enabled in the document.
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ExportedWithLineNumbers.txt");

        // Save the document as plain text using the configured options.
        doc.Save(outputPath, saveOptions);

        // Display the resulting file content on the console.
        Console.WriteLine("Exported text with line numbers:");
        Console.WriteLine(File.ReadAllText(outputPath));
    }
}
