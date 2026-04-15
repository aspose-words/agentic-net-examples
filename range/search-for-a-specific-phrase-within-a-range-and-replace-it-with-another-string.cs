using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text that contains the phrase we want to replace.
        builder.Writeln("This is a sample document. The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Hello World! This phrase will be replaced.");

        // Define the phrase to search for and the replacement text.
        string searchPhrase = "Hello World";
        string replacementText = "Greetings from Aspose.Words";

        // Perform a simple find-and-replace on the whole document range.
        int replacementsMade = doc.Range.Replace(searchPhrase, replacementText);

        // Optional: output the number of replacements to the console.
        Console.WriteLine($"Replacements performed: {replacementsMade}");

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "ModifiedDocument.docx");
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
