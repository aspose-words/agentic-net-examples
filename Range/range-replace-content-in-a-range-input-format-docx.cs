using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class RangeReplaceExample
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Define the text to find and its replacement.
        string pattern = "_FullName_";
        string replacement = "John Doe";

        // Perform a simple find-and-replace on the whole document range.
        // The Replace method returns the number of replacements made.
        int count = doc.Range.Replace(pattern, replacement);

        Console.WriteLine($"Replacements made: {count}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
