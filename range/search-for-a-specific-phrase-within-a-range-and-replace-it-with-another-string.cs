using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs that contain the phrase we want to replace.
        builder.Writeln("Hello Aspose.Words! This is a sample document.");
        builder.Writeln("We will replace the phrase 'sample document' with 'demo file'.");

        // Define the phrase to search for and its replacement.
        string searchPhrase = "sample document";
        string replacement = "demo file";

        // Perform a simple find-and-replace on the whole‑document range.
        int replacementsMade = doc.Range.Replace(searchPhrase, replacement);

        // Output the number of replacements (optional verification).
        Console.WriteLine($"Replacements made: {replacementsMade}");

        // Save the modified document to the local file system.
        doc.Save("Output.docx");
    }
}
