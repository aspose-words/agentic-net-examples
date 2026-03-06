using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Set up find/replace options.
        // In this example we also change the paragraph alignment of any paragraph that contains a match.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Right;

        // Perform the find-and-replace operation on the whole document range.
        // Replace every occurrence of the placeholder "{Name}" with "John Doe".
        doc.Range.Replace("{Name}", "John Doe", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
