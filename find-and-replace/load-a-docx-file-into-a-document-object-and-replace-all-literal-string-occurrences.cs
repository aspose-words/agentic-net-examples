using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly
using Newtonsoft.Json; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "input.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.docx");

        // -----------------------------------------------------------------
        // Create a sample DOCX file with known text to be replaced.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("Replace the word TARGET wherever it appears.");
        builder.Writeln("TARGET appears multiple times: TARGET, TARGET.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Perform a literal string find-and-replace.
        // -----------------------------------------------------------------
        string findText = "TARGET";
        string replaceText = "REPLACED";
        FindReplaceOptions options = new FindReplaceOptions(); // default options
        int replacedCount = loadedDoc.Range.Replace(findText, replaceText, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException($"No occurrences of \"{findText}\" were found to replace.");
        }

        // -----------------------------------------------------------------
        // Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);

        // Optional: indicate success (no interactive input required).
        Console.WriteLine($"Replaced {replacedCount} occurrence(s). Output saved to: {outputPath}");
    }
}
