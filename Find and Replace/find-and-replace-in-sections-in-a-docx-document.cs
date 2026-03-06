using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Replace throughout the entire document (all sections).
        int totalReplacements = doc.Range.Replace(findText, replaceText);
        Console.WriteLine($"Total replacements made: {totalReplacements}");

        // Example: replace only within a specific section (e.g., the second section).
        if (doc.Sections.Count > 1)
        {
            Section secondSection = doc.Sections[1];
            int sectionReplacements = secondSection.Range.Replace(findText, replaceText);
            Console.WriteLine($"Replacements in second section: {sectionReplacements}");
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
