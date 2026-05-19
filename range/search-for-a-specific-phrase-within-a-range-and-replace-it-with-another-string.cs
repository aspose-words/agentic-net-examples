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
        builder.Writeln("Aspose.Words makes document processing easy. Replace the phrase \"lazy dog\" now.");

        // Define the phrase to search for and the replacement text.
        string searchPhrase = "lazy dog";
        string replacementText = "energetic cat";

        // Perform a simple find-and-replace on the whole document range.
        int replacementsMade = doc.Range.Replace(searchPhrase, replacementText);

        // Optional verification (can be removed if not needed).
        Console.WriteLine($"Replacements performed: {replacementsMade}");
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(doc.GetText().Trim());

        // Save the modified document to the local file system.
        string outputPath = "ModifiedDocument.docx";
        doc.Save(outputPath);
    }
}
