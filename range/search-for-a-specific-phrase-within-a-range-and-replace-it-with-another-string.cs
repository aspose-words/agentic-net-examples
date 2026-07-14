using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text that contains the phrase we want to replace.
        builder.Writeln("This is a sample document. The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Find and replace the phrase \"lazy dog\" with \"energetic cat\".");

        // Define the phrase to search for and its replacement.
        string searchPhrase = "lazy dog";
        string replacement = "energetic cat";

        // Perform the replacement on the whole-document range.
        int replacementsMade = doc.Range.Replace(searchPhrase, replacement);

        // Output the result of the operation.
        Console.WriteLine($"Replacements made: {replacementsMade}");
        Console.WriteLine("Updated document text:");
        Console.WriteLine(doc.GetText().Trim());

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
