using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the plain‑text file to be loaded.
        string txtPath = "input.txt";

        // Configure load options to recognize list items that use whitespace as a delimiter.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // Load the plain‑text file into a Document using the constructor that accepts a file name and load options.
        Document doc = new Document(txtPath, loadOptions);

        // Output the number of lists that Aspose.Words detected.
        Console.WriteLine($"Lists detected: {doc.Lists.Count}");

        // Iterate through paragraphs and print those that were recognized as list items.
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            if (para.IsListItem)
                Console.WriteLine($"List item: {para.GetText().TrimEnd()}");
        }
    }
}
