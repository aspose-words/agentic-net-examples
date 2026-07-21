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

        // Insert sample text into the document.
        builder.Writeln("Hello world! This is a sample document.");

        // Replace the word "Hello" with "Hi" using the document's range.
        int replacementsMade = doc.Range.Replace("Hello", "Hi");

        // (Optional) Verify that a replacement occurred.
        // Console.WriteLine($"Replacements made: {replacementsMade}");

        // Save the modified document to the local file system.
        doc.Save("Output.docx");
    }
}
