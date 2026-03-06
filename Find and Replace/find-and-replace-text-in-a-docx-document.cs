using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindAndReplaceExample
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Path where the modified DOCX will be saved.
        string outputPath = @"C:\Docs\output.docx";

        // Text to find and its replacement.
        string pattern = "_FullName_";
        string replacement = "John Doe";

        // Load the document from disk.
        Document doc = new Document(inputPath);

        // Perform a simple find-and-replace operation on the whole document.
        // The Replace method returns the number of replacements made.
        int replacementCount = doc.Range.Replace(pattern, replacement);

        // Optionally, you can inspect the result.
        Console.WriteLine($"Number of replacements made: {replacementCount}");
        Console.WriteLine("Resulting document text (trimmed):");
        Console.WriteLine(doc.GetText().Trim());

        // Save the modified document.
        doc.Save(outputPath);
    }
}
