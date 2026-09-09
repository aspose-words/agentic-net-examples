using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing the word to be replaced.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("The quick brown fox is quick and clever.");
        builder.Writeln("Quickness is a virtue.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Perform a find-and-replace operation and capture the number of replacements.
        const string findText = "quick";
        const string replaceText = "swift";
        int replacementCount = loaded.Range.Replace(findText, replaceText, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Output the count of replacements performed.
        Console.WriteLine($"Number of replacements performed: {replacementCount}");
    }
}
