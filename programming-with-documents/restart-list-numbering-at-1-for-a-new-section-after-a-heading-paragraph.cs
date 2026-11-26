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

        // First heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Section 1");

        // First list (normal numbering).
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers(); // End the list.

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Section 2");

        // Create a list that restarts numbering at each new section.
        List restartList = doc.Lists.Add(ListTemplate.NumberDefault);
        restartList.IsRestartAtEachSection = true;

        // Apply the restart list.
        builder.ListFormat.List = restartList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers(); // End the list.

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
