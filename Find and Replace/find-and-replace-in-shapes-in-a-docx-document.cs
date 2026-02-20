using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace to process text inside shapes.
        FindReplaceOptions options = new FindReplaceOptions();
        options.IgnoreShapes = false; // include shape text in the search.

        // Example: replace a literal string inside the document and its shapes.
        doc.Range.Replace("PLACEHOLDER", "Actual Value", options);

        // Example: replace using a regular expression inside the document and its shapes.
        // doc.Range.Replace(new Regex(@"\bOldWord\b"), "NewWord", options);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
