using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text that contains the placeholder we will replace.
        builder.Writeln("Hello _FullName_, welcome to the system.");
        builder.Writeln("Your username is _FullName_.");

        // Perform the replace operation.
        // The Replace method returns the number of replacements made.
        int replacementCount = doc.Range.Replace("_FullName_", "John Doe");

        // Output the number of replacements.
        Console.WriteLine($"Number of replacements performed: {replacementCount}");

        // Save the modified document (optional, demonstrates lifecycle usage).
        doc.Save("Result.docx");
    }
}
