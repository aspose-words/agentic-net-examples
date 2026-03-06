using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Define the placeholder text that exists in the document
        // and the dynamic value that should replace it.
        string placeholder = "_FullName_";
        string replacement = "John Doe";

        // Perform a find‑and‑replace on the whole document range.
        // This overload uses a case‑insensitive search and replaces all occurrences.
        int replacedCount = doc.Range.Replace(placeholder, replacement);

        Console.WriteLine($"Number of replacements performed: {replacedCount}");

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
