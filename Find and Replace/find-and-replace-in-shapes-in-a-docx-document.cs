using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Set up find/replace options.
        // By default IgnoreShapes = false, which means the operation will also search inside shapes.
        // Explicitly set it for clarity.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreShapes = false   // Do NOT ignore shapes – replace text inside them.
        };

        // Perform the replacement across the whole document range.
        // Example: replace the placeholder "[NAME]" with "John Doe".
        int replacementsMade = doc.Range.Replace("[NAME]", "John Doe", options);

        Console.WriteLine($"Replacements performed: {replacementsMade}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
