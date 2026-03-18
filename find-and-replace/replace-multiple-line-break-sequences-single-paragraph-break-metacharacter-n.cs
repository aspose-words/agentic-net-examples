using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceMultipleLineBreaks
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Define a regular expression that matches two or more consecutive line break characters
        // (\r\n, \r, or \n). This captures any sequence of blank lines.
        Regex multipleLineBreaks = new Regex(@"(\r\n|\r|\n){2,}", RegexOptions.Multiline);

        // Replace each match with a single line feed character (\n).
        // In Aspose.Words the replacement string is processed for control characters,
        // so "\n" will be interpreted as a line feed (paragraph break when saved as text).
        doc.Range.Replace(multipleLineBreaks, "\n");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
