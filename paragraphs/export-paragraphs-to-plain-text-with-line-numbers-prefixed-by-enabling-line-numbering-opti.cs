using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs.
        builder.Writeln("First paragraph of the document.");
        builder.Writeln("Second paragraph follows the first one.");
        builder.Writeln("Third paragraph is here to demonstrate line numbering.");

        // Enable line numbering for the first section (this affects layout, not plain‑text export).
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        pageSetup.LineStartingNumber = 1;               // Start numbering at 1.
        pageSetup.LineNumberCountBy = 1;                // Number every line.
        pageSetup.LineNumberRestartMode = LineNumberRestartMode.Continuous;
        pageSetup.LineNumberDistanceFromText = 0;       // Default distance.

        // Prepare the output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Export the document to plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            // No ExportLineNumbers property in this version; we will add numbers manually.
            ParagraphBreak = Environment.NewLine
        };

        // Save the raw text first (optional, can be omitted if only the numbered version is needed).
        string rawTxtPath = Path.Combine(outputDir, "DocumentRaw.txt");
        doc.Save(rawTxtPath, saveOptions);

        // Read the raw text, prefix each line with its line number, and write the final file.
        string rawText = File.ReadAllText(rawTxtPath);
        string[] lines = rawText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

        StringBuilder numberedBuilder = new StringBuilder();
        for (int i = 0; i < lines.Length; i++)
        {
            // Prefix with line number (starting at 1) and a space.
            numberedBuilder.AppendLine($"{i + 1} {lines[i]}");
        }

        string numberedTxtPath = Path.Combine(outputDir, "DocumentWithLineNumbers.txt");
        File.WriteAllText(numberedTxtPath, numberedBuilder.ToString());

        // Display the resulting text in the console.
        Console.WriteLine("Document saved with line numbers:");
        Console.WriteLine(File.ReadAllText(numberedTxtPath));
    }
}
