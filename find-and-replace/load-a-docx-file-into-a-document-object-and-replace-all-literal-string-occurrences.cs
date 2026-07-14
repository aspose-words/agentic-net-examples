using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "input.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.docx");

        // -----------------------------------------------------------------
        // Create a sample DOCX file containing the text to be replaced.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("Replace the word old with new.");
        builder.Writeln("Another old occurrence here.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document, perform a literal string replacement, and save.
        // -----------------------------------------------------------------
        var loadedDoc = new Document(inputPath);
        int replacedCount = loadedDoc.Range.Replace("old", "new", new FindReplaceOptions());

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        loadedDoc.Save(outputPath);
    }
}
