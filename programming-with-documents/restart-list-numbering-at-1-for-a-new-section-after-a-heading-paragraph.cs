using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace RestartListNumberingExample
{
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

            // First numbered list.
            // Use the default numbered list template.
            List firstList = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = firstList;
            builder.ListFormat.ListLevelNumber = 0; // top level
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers(); // End of the first list.

            // Second heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Section 2");

            // Restart numbering at 1 for the new list.
            // Create a new list and set its start number to 1.
            List secondList = doc.Lists.Add(ListTemplate.NumberDefault);
            secondList.ListLevels[0].StartAt = 1; // restart numbering
            builder.ListFormat.List = secondList;
            builder.ListFormat.ListLevelNumber = 0; // top level
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers(); // End of the second list.

            // Save the document.
            const string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
