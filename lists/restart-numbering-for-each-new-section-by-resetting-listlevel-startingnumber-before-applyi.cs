using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on a built‑in template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Optional: let Word know the list should restart at each section.
        numberedList.IsRestartAtEachSection = true;

        // First section.
        builder.Writeln("Section 1");
        // Reset the starting number for the first level before using the list.
        numberedList.ListLevels[0].StartAt = 1;
        builder.ListFormat.List = numberedList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.ListFormat.RemoveNumbers();

        // Insert a section break (new page) to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section.
        builder.Writeln("Section 2");
        // Reset the starting number again for the new section.
        numberedList.ListLevels[0].StartAt = 1;
        builder.ListFormat.List = numberedList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output file.
        doc.Save("RestartNumberingPerSection.docx");
    }
}
