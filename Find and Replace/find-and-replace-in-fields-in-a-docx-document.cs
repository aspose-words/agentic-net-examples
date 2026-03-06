using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options to process text inside fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Set to false (default) to include fields in the search.
            // Change to true if you want to ignore field contents.
            IgnoreFields = false
        };

        // Perform the replacement across the whole document range.
        // Example: replace the placeholder "_FullName_" with "John Doe".
        int replacementsMade = doc.Range.Replace("_FullName_", "John Doe", options);

        // Optionally, output the number of replacements.
        Console.WriteLine($"Replacements performed: {replacementsMade}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
