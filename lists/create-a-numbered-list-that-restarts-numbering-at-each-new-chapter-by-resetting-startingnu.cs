using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on the default template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // First chapter.
        InsertChapter(builder, "Chapter 1");
        // Reset the list numbering to start from 1 for this chapter.
        numberedList.ListLevels[0].StartAt = 1;
        // Apply the list to the builder and add some items.
        builder.ListFormat.List = numberedList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        // End the list for this chapter.
        builder.ListFormat.RemoveNumbers();

        // Insert a section break to start a new chapter on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second chapter.
        InsertChapter(builder, "Chapter 2");
        // Reset the list numbering again.
        numberedList.ListLevels[0].StartAt = 1;
        builder.ListFormat.List = numberedList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "NumberedListRestart.docx");
        doc.Save(outputPath);
    }

    // Helper method to insert a chapter heading.
    private static void InsertChapter(DocumentBuilder builder, string title)
    {
        // Use Heading 1 style for chapter titles.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln(title);
        // Return to normal style for list items.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
    }
}
