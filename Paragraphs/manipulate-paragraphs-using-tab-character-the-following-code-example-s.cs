using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure every list level to use a TAB character as the separator
        // between the list label (e.g., "1.") and the paragraph text.
        for (int i = 0; i < list.ListLevels.Count; i++)
        {
            list.ListLevels[i].TrailingCharacter = ListTrailingCharacter.Tab;
        }

        // Apply the list to the builder so that subsequent paragraphs become list items.
        builder.ListFormat.List = list;

        // First level item.
        builder.Writeln("Item 1");

        // Increase the list level (second level).
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 2");

        // Increase the list level again (third level).
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 3");

        // Remove list formatting from the builder cursor.
        builder.ListFormat.RemoveNumbers();

        // Save the document as a DOCX file (preserves the list formatting).
        doc.Save("ListWithTabs.docx");

        // When exporting to plain text we want the indentation to be a TAB character.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.ListIndentation.Character = '\t'; // Use TAB for indentation.
        txtOptions.ListIndentation.Count = 1;        // One TAB per list level.

        // Save the same document as plain text using the configured options.
        doc.Save("ListWithTabs.txt", txtOptions);
    }
}
