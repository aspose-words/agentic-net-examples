using System;
using System.Text.RegularExpressions;
using Aspose.Words;

class ReplaceMultipleSpaces
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Replace any occurrence of two or more consecutive space characters with a single space.
        // The regular expression "[ ]{2,}" matches a sequence of at least two spaces.
        doc.Range.Replace(new Regex(@"[ ]{2,}"), " ");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
