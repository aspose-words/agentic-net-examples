using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class ExportParagraphsWithLineNumbers
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable line numbering for the first section.
        // Numbers will start at 1, appear on every line, and restart on each new page.
        PageSetup pageSetup = builder.PageSetup;
        pageSetup.LineStartingNumber = 1;               // First line number.
        pageSetup.LineNumberCountBy = 1;                // Number every line.
        pageSetup.LineNumberRestartMode = LineNumberRestartMode.RestartPage; // Restart each page.
        pageSetup.LineNumberDistanceFromText = 30.0;    // Distance from the text (points).

        // Add several paragraphs to demonstrate line numbering.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph with a bit more text to wrap onto the next line.");
        builder.Writeln("Third paragraph.");
        builder.Writeln("Fourth paragraph.");

        // Configure TxtSaveOptions – no special settings required for line numbers.
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Define output path (relative to the executable's working directory).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ParagraphsWithLineNumbers.txt");

        // Save the document as plain text; line numbers will be prefixed automatically.
        doc.Save(outputPath, saveOptions);

        // Inform the user where the file was saved.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
