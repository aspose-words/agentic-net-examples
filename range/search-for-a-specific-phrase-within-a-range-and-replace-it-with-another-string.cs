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

        // Add some text that contains the phrase we will search for.
        builder.Writeln("This document contains a sample phrase that will be replaced.");

        // Define the phrase to find and the replacement text.
        string phraseToFind = "sample phrase";
        string replacementText = "replaced text";

        // Perform a simple find-and-replace on the whole‑document range.
        int replacementsMade = doc.Range.Replace(phraseToFind, replacementText);

        // Output the result of the operation.
        Console.WriteLine($"Replacements made: {replacementsMade}");
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(doc.GetText().Trim());

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
