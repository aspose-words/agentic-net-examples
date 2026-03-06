using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Regular expression to find all sequences of digits.
        Regex pattern = new Regex(@"\d+");

        // Replace each match with a paragraph break.
        // The special meta‑character "&p" inserts a paragraph break.
        string replacement = "&p";

        // Execute the find‑and‑replace operation.
        int replacedCount = doc.Range.Replace(pattern, replacement);

        Console.WriteLine($"Number of replacements performed: {replacedCount}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
