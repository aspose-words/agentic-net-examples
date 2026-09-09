using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for inserting content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        // Optional: let the list restart automatically at each new section.
        list.IsRestartAtEachSection = true;

        // ---------- First section ----------
        // Apply the list to the first two paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Section 1 – Item 1");
        builder.Writeln("Section 1 – Item 2");
        // End the list for this section.
        builder.ListFormat.RemoveNumbers();

        // Insert a section break (new page) to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Second section ----------
        // Reset the starting number of the first list level before reusing the list.
        list.ListLevels[0].StartAt = 1;

        // Apply the same list to the new section.
        builder.ListFormat.List = list;
        builder.Writeln("Section 2 – Item 1");
        builder.Writeln("Section 2 – Item 2");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "RestartListPerSection.docx");
        doc.Save(outputPath);
    }
}
