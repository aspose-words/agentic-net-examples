using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class InsertRestartNumTag
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Define a regular expression that matches the start of a numbered paragraph.
        // This pattern captures a sequence of digits followed by a period (e.g., "1.", "23.", etc.).
        Regex numberedParagraphPattern = new Regex(@"^(\d+\.)", RegexOptions.Multiline);

        // Set up find‑replace options. No special flags are required for this operation.
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace the matched number with itself followed by the <restartNum/> tag.
        // The replacement string uses "$1" to insert the captured number and then adds the tag.
        doc.Range.Replace(numberedParagraphPattern, "$1<restartNum/>", options);

        // Save the modified document in DOCX format.
        doc.Save("Output.docx");
    }
}
