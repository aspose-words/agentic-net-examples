using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Jackson will meet you in Jacksonville.");

        // Save the sample document to the local file system.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just saved.
        Document loadedDoc = new Document(inputPath);

        // Configure find-and-replace to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace("Jackson", "Louis", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Output the resulting text to the console (optional verification).
        Console.WriteLine("Replacement count: " + replacedCount);
        Console.WriteLine("Resulting text:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
