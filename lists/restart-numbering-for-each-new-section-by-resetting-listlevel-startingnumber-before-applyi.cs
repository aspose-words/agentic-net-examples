using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class RestartListNumbering
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // First section – the list starts at 1.
        builder.Writeln("Section 1:");
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.ListFormat.RemoveNumbers();

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Reset the starting number of the first list level before reusing the list.
        list.ListLevels[0].StartAt = 1;

        // Second section – numbering restarts from 1.
        builder.Writeln("Section 2:");
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "RestartListNumbering.docx");
        doc.Save(outputPath);
    }
}
