using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindAndReplaceExample
{
    static void Main()
    {
        // Define input and output file paths.
        string inputPath = @"C:\Data\Input.docx";
        string outputPath = @"C:\Data\Output.docx";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder for possible additional editing (not required for replace itself).
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example: replace a placeholder token with actual text.
        string placeholder = "_FullName_";
        string replacement = "John Doe";

        // Perform a simple case‑insensitive find‑and‑replace across the whole document.
        int replacedCount = doc.Range.Replace(placeholder, replacement);

        // Optionally, use FindReplaceOptions for more control (e.g., case‑sensitive, whole word only).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,               // ignore case
            FindWholeWordsOnly = true        // replace only whole word matches
        };
        // Replace another token using the options.
        doc.Range.Replace("[Date]", DateTime.Today.ToString("MMMM d, yyyy"), options);

        // Save the modified document.
        doc.Save(outputPath);
    }
}
