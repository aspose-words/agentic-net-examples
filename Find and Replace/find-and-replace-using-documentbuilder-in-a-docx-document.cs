using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class FindReplaceExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello _UserName_, welcome to the Aspose.Words tutorial.");
        builder.Writeln("Please replace _UserName_ with your actual name.");

        // Define the text to find and its replacement.
        string findText = "_UserName_";
        string replaceText = "John Doe";

        // Simple case‑insensitive replace without extra options.
        int count = doc.Range.Replace(findText, replaceText);
        Console.WriteLine($"Replacements made: {count}");

        // If you need more control (e.g., match whole words only), use FindReplaceOptions.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,               // ignore case
            FindWholeWordsOnly = false       // replace even if part of a larger word
        };
        // Example of a second replace using the options object.
        doc.Range.Replace("Aspose.Words", "Aspose.Words (v2)", options);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
