using System;
using System.Text.RegularExpressions;
using Aspose.Words;

class ReplaceDoubleSpaces
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Replace any sequence of two or more spaces with a single space.
        doc.Range.Replace(new Regex(@" {2,}"), " ");

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
