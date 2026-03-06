using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;

class RestartListNumbering
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a default numbered list to the document's list collection.
        doc.Lists.Add(ListTemplate.NumberDefault);

        // Retrieve the first list (the one we just added).
        List list = doc.Lists[0];

        // Enable restarting of the list at each section (DOC format).
        list.IsRestartAtEachSection = true;

        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = list;

        // First section items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break; numbering will restart after this break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section items – numbering starts again from 1.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Remove list formatting from further paragraphs (optional).
        builder.ListFormat.RemoveNumbers();

        // Save the document in DOC format.
        doc.Save("RestartList.doc", SaveFormat.Doc);
    }
}
