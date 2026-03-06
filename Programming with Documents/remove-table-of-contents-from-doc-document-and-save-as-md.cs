using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class RemoveTocAndSaveAsMarkdown
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.docx");

        // Iterate over all fields in reverse order and remove any Table of Contents (TOC) fields.
        for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
        {
            Field field = doc.Range.Fields[i];
            if (field.Type == FieldType.FieldTOC)
                field.Remove(); // Removes the TOC field from the document.
        }

        // Prepare Markdown save options (default settings are sufficient for this task).
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Save the modified document as a Markdown file.
        doc.Save("OutputDocument.md", saveOptions);
    }
}
