using System;
using Aspose.Words;
using Aspose.Words.Replacing;
// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Words.Range;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample text that contains the phrase we want to replace.
        builder.Writeln("Hello world! This is a test. Hello world!");

        // Define the range we want to operate on – in this case the first paragraph.
        AsposeRange paragraphRange = doc.FirstSection.Body.Paragraphs[0].Range;

        // Search for the phrase "Hello" and replace it with "Hi".
        // The Replace method returns the number of replacements made.
        int replacementCount = paragraphRange.Replace("Hello", "Hi");

        // Output the result to the console.
        Console.WriteLine($"Number of replacements performed: {replacementCount}");
        Console.WriteLine("Modified document text:");
        Console.WriteLine(doc.GetText().Trim());

        // Save the modified document to a file.
        doc.Save("ModifiedDocument.docx");
    }
}
