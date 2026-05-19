using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for the sample files.
        string workDir = Path.Combine(Environment.CurrentDirectory, "FindReplaceDemo");
        Directory.CreateDirectory(workDir);

        // Create a sample document with several paragraphs containing the word "TARGET".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with TARGET.");
        builder.Writeln("Second paragraph without the keyword.");
        builder.Writeln("Third paragraph with TARGET again.");
        string inputPath = Path.Combine(workDir, "input.docx");
        doc.Save(inputPath);

        // Load the document we just created.
        Document loadedDoc = new Document(inputPath);

        // Replace each occurrence of "TARGET" with "TARGET" followed by a paragraph break meta‑character.
        // The meta‑character "&p" tells Aspose.Words to insert a paragraph break in the replacement text.
        int replacedCount = loadedDoc.Range.Replace("TARGET", "TARGET&p", new FindReplaceOptions());

        // Verify that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        string outputPath = Path.Combine(workDir, "output.docx");
        loadedDoc.Save(outputPath);

        // Optional: output a simple confirmation (no interactive input required).
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
