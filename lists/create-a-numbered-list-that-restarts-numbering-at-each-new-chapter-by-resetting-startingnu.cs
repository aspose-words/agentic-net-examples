using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListRestartExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define a base numbered list template.
            List baseList = doc.Lists.Add(ListTemplate.NumberDefault);

            // -------------------- Chapter 1 --------------------
            // Insert a heading for the chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            // Create a copy of the base list and reset its starting number.
            List chapter1List = doc.Lists.AddCopy(baseList);
            chapter1List.ListLevels[0].StartAt = 1; // Restart numbering at 1.

            // Apply the list to the following paragraphs.
            builder.ListFormat.List = chapter1List;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.Writeln("Item 3");
            builder.ListFormat.RemoveNumbers(); // End the list.

            // -------------------- Chapter 2 --------------------
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");

            // Create another list copy for the second chapter and reset its start.
            List chapter2List = doc.Lists.AddCopy(baseList);
            chapter2List.ListLevels[0].StartAt = 1; // Restart numbering at 1.

            builder.ListFormat.List = chapter2List;
            builder.Writeln("Item A");
            builder.Writeln("Item B");
            builder.Writeln("Item C");
            builder.ListFormat.RemoveNumbers();

            // Save the document to disk.
            doc.Save("NumberedListByChapter.docx");
        }
    }
}
