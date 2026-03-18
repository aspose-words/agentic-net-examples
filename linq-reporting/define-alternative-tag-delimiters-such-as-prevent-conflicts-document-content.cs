using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class AlternativeTagDelimiters
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Replace the opening mustache delimiter "{{" with the alternative "[[".
        // Use FindReplaceOptions to treat the replacement as plain text.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Ensure that structured document tags are not split during replacement.
            IgnoreStructuredDocumentTags = true
        };
        doc.Range.Replace("{{", "[[", options);

        // Replace the closing mustache delimiter "}}" with the alternative "]]".
        doc.Range.Replace("}}", "]]", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
