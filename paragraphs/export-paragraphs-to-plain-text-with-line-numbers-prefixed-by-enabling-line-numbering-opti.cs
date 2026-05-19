using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to add sample paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Configure plain‑text save options.
        TxtSaveOptions txtOptions = new TxtSaveOptions();

        // Save the document as plain text.
        string tempFile = "TempExport.txt";
        doc.Save(tempFile, txtOptions);

        // Read the exported text, prefix each line with its 1‑based line number, and display the result.
        string[] lines = File.ReadAllLines(tempFile);
        for (int i = 0; i < lines.Length; i++)
        {
            Console.WriteLine($"{i + 1}: {lines[i]}");
        }

        // Write the numbered text to a final file.
        string outputFile = "ExportedWithLineNumbers.txt";
        string[] numberedLines = new string[lines.Length];
        for (int i = 0; i < lines.Length; i++)
        {
            numberedLines[i] = $"{i + 1}: {lines[i]}";
        }
        File.WriteAllLines(outputFile, numberedLines);
    }
}
