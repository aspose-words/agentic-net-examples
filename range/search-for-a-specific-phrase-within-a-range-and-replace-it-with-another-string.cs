using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing the target phrase.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document. Hello World! This phrase will be replaced.");

        // Save the original document.
        doc.Save("Original.docx");

        // Define the phrase to search for and its replacement.
        string searchPhrase = "Hello World";
        string replaceWith = "Hi Universe";

        // Perform the replacement on the whole-document range.
        doc.Range.Replace(searchPhrase, replaceWith, new FindReplaceOptions());

        // Save the modified document.
        doc.Save("Modified.docx");

        // Output the resulting text to verify the replacement.
        Console.WriteLine("Replacement performed. Modified document text:");
        Console.WriteLine(doc.Range.Text);
    }
}
