using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceWithParagraph
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create a FindReplaceOptions instance to customize the replace operation.
        FindReplaceOptions options = new FindReplaceOptions();

        // Example: set paragraph formatting for the newly inserted paragraph.
        // This will apply the specified alignment to the paragraph that is inserted.
        options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Define the text to find and the replacement text.
        // The "&p" meta‑character inserts a paragraph break after the replacement text,
        // effectively creating a new paragraph.
        string findText = "[PLACEHOLDER]";
        string replaceText = "This is the new paragraph content.&p";

        // Perform the find‑and‑replace operation on the whole document range.
        doc.Range.Replace(findText, replaceText, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
