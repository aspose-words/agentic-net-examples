using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Text to search for and its replacement.
        string findText = "Placeholder";
        string replaceText = "Actual Value";

        // Configure find/replace to ignore the internal structure of StructuredDocumentTags.
        // This treats the content of each SDT as plain text, allowing replacements that span across tags.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreStructuredDocumentTags = true
        };

        // Perform the find-and-replace operation on the whole document.
        doc.Range.Replace(findText, replaceText, options);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
