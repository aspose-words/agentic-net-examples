using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First chapter.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Create a numbered list for the first chapter.
        List list1 = doc.Lists.Add(ListTemplate.NumberDefault);
        // Ensure numbering starts at 1.
        list1.ListLevels[0].StartAt = 1;

        // Apply the list to the following paragraphs.
        builder.ListFormat.List = list1;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First item of chapter 1");
        builder.Writeln("Second item of chapter 1");
        builder.Writeln("Third item of chapter 1");
        builder.ListFormat.RemoveNumbers();

        // Second chapter.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Create a new numbered list for the second chapter (restarts numbering).
        List list2 = doc.Lists.Add(ListTemplate.NumberDefault);
        list2.ListLevels[0].StartAt = 1; // Reset starting number.

        builder.ListFormat.List = list2;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First item of chapter 2");
        builder.Writeln("Second item of chapter 2");
        builder.Writeln("Third item of chapter 2");
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("NumberedList.docx");
    }
}
