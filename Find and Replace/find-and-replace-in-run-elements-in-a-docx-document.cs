using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class RunReplaceExample
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document(@"C:\Docs\Input.docx");

        // Text to find and its replacement.
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Simple replace across the whole document (includes runs).
        // The Range.Replace method processes the document's runs internally.
        int replacements = doc.Range.Replace(findText, replaceText);

        // Output the number of replacements made (optional).
        Console.WriteLine($"Replacements performed: {replacements}");

        // Save the modified document.
        doc.Save(@"C:\Docs\Output.docx");
    }
}
