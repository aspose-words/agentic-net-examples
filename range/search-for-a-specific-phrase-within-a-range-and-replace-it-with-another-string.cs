using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with the target phrase.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("Hello World! This phrase will be replaced.");
        builder.Writeln("End of document.");

        // Save the original document (optional, just to demonstrate file creation).
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document from the file to ensure the load workflow is exercised.
        Document loadedDoc = new Document(originalPath);

        // Define the phrase to search for and its replacement.
        string searchPhrase = "Hello World";
        string replacement = "Hi Universe";

        // Perform the replacement on the whole document range.
        FindReplaceOptions options = new FindReplaceOptions();
        loadedDoc.Range.Replace(searchPhrase, replacement, options);

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        loadedDoc.Save(modifiedPath);

        // Verify the replacement by outputting the document's plain text to the console.
        string resultText = loadedDoc.Range.Text;
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(resultText);
    }
}
