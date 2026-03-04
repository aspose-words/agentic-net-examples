using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        // (Assumes the file "Input.docx" exists in the same folder as the executable.)
        Document doc = new Document("Input.docx");

        // Define the text to find and its replacement.
        string pattern = "_FullName_";
        string replacement = "John Doe";

        // Perform a simple find-and-replace across the whole document.
        // This uses the Range.Replace(string, string) overload.
        int count = doc.Range.Replace(pattern, replacement);

        // Optionally, you can use FindReplaceOptions for more control.
        // Example: make the replace operation case‑insensitive.
        // FindReplaceOptions options = new FindReplaceOptions { MatchCase = false };
        // int count = doc.Range.Replace(pattern, replacement, options);

        Console.WriteLine($"Number of replacements made: {count}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
