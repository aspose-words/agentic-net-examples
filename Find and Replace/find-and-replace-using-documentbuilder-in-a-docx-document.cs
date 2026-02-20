using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk.
        Document doc = new Document("Input.docx");

        // Attach a DocumentBuilder to the loaded document (optional, useful for further edits).
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Perform a find-and-replace operation.
        // This replaces every occurrence of the placeholder text "[PLACEHOLDER]" with "Hello World".
        FindReplaceOptions options = new FindReplaceOptions(); // default options
        doc.Range.Replace("[PLACEHOLDER]", "Hello World", options);

        // Save the modified document to a new file.
        doc.Save("Output.docx");
    }
}
