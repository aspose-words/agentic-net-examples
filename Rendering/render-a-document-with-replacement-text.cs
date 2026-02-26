using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a placeholder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Greetings, _FullName_!");

        // Perform a find-and-replace operation on the whole document.
        // Replace the placeholder with the actual name.
        int replacements = doc.Range.Replace("_FullName_", "John Doe");

        // Optional: verify that a replacement was made.
        Console.WriteLine($"Number of replacements performed: {replacements}");

        // Save the resulting document to disk.
        doc.Save("Result.docx");
    }
}
