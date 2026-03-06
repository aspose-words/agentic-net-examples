using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceInShapes
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\Input.docx";

        // Path where the modified DOCX will be saved.
        string outputPath = @"C:\Docs\Output.docx";

        // Load the document.
        Document doc = new Document(inputPath);

        // Configure find/replace to include text that resides inside shapes.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // By default this property is false, but we set it explicitly for clarity.
            IgnoreShapes = false
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("OldText", "NewText", options);

        // Output the number of replacements made.
        Console.WriteLine($"Replacements performed: {replacedCount}");

        // Save the updated document.
        doc.Save(outputPath);
    }
}
