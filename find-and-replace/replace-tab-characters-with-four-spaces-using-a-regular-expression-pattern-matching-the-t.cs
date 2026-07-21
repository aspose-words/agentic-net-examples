using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Required package reference (not used directly in this example)

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // Create a sample document that contains tab characters.
        // -----------------------------------------------------------------
        Document doc = new Document();                     // Create a new blank document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Column1\tColumn2\tColumn3");      // Write a line with tabs.
        doc.Save(inputPath);                               // Save the source document.

        // -----------------------------------------------------------------
        // Load the document and replace each tab with four spaces.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);         // Load the previously saved document.

        // Regular expression that matches a tab character.
        Regex tabRegex = new Regex(@"\t");

        // Replacement string: four space characters.
        const string fourSpaces = "    ";

        // FindReplaceOptions can be customized if needed.
        FindReplaceOptions options = new FindReplaceOptions();

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(tabRegex, fourSpaces, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one tab character to be replaced.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Indicate success.
        Console.WriteLine($"Replaced {replacedCount} tab character(s). Output saved to: {outputPath}");
    }
}
