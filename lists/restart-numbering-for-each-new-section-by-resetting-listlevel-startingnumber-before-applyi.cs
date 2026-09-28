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

        // Create two sections, each with its own numbered list that restarts from 1.
        for (int sectionIndex = 1; sectionIndex <= 2; sectionIndex++)
        {
            // Add a heading for the section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Section {sectionIndex}");
            builder.ParagraphFormat.ClearFormatting();

            // Create a new numbered list and set its starting number to 1.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);
            list.ListLevels[0].StartAt = 1;   // Reset numbering for this list.

            // Add three items to the list.
            for (int itemIndex = 1; itemIndex <= 3; itemIndex++)
            {
                builder.ListFormat.List = list;
                builder.Writeln($"Item {itemIndex} in Section {sectionIndex}");
            }

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Insert a blank line between sections.
            builder.Writeln();
        }

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
