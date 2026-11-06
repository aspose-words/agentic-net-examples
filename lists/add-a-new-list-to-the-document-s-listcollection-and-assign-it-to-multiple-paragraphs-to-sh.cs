using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new numbered list and obtain its reference.
        List sharedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the first paragraph.
        builder.ListFormat.List = sharedList;
        builder.Writeln("First item of the shared list.");

        // Apply the same list to the second paragraph.
        builder.Writeln("Second item of the shared list.");

        // End the list formatting before adding normal text.
        builder.ListFormat.RemoveNumbers();
        builder.Writeln("This paragraph is not part of the list.");

        // Validation: ensure exactly two paragraphs are list items.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Paragraph.ListFormat provides list information.
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        if (listParagraphCount != 2)
            throw new InvalidOperationException($"Expected 2 list paragraphs, but found {listParagraphCount}.");

        // Save the document to a file.
        doc.Save("SharedListExample.docx");
    }
}
