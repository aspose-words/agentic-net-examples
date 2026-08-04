using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Ensure the first level starts at 1 for the first section.
        list.ListLevels[0].StartAt = 1;

        // Apply the list to the first section.
        builder.ListFormat.List = list;
        builder.Writeln("Section 1 - Item 1");
        builder.Writeln("Section 1 - Item 2");

        // Insert a section break (new page) to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Reset the starting number for the next section.
        list.ListLevels[0].StartAt = 1;

        // Apply the same list to the second section.
        builder.ListFormat.List = list;
        builder.Writeln("Section 2 - Item 1");
        builder.Writeln("Section 2 - Item 2");

        // Save the document to disk.
        doc.Save("RestartListPerSection.docx");
    }
}
