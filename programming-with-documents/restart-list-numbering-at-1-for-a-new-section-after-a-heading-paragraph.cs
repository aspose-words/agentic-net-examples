using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;

public class RestartListNumbering
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list and enable restarting at each section.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        numberedList.IsRestartAtEachSection = true;

        // Apply the list to the builder.
        builder.ListFormat.List = numberedList;

        // First heading (Section 1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Section 1");

        // List items for the first section.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second heading (Section 2).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Section 2");

        // List items for the second section – numbering restarts at 1.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local file system.
        doc.Save("RestartList.docx");
    }
}
