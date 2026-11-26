using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace NumberedListRestartExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------- Chapter 1 --------------------
            builder.Writeln("Chapter 1");

            // Create a numbered list for Chapter 1.
            List chapter1List = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = chapter1List;
            builder.ListFormat.ListLevelNumber = 0; // Top level.

            // Add list items.
            builder.Writeln("First item of Chapter 1");
            builder.Writeln("Second item of Chapter 1");
            builder.Writeln("Third item of Chapter 1");

            // End of Chapter 1 list.
            builder.ListFormat.RemoveNumbers();

            // Add some spacing between chapters.
            builder.Writeln();

            // -------------------- Chapter 2 --------------------
            builder.Writeln("Chapter 2");

            // Create a new numbered list for Chapter 2 to restart numbering.
            List chapter2List = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = chapter2List;
            builder.ListFormat.ListLevelNumber = 0; // Top level.

            // Add list items.
            builder.Writeln("First item of Chapter 2");
            builder.Writeln("Second item of Chapter 2");

            // End of Chapter 2 list.
            builder.ListFormat.RemoveNumbers();

            // -------------------- Validation --------------------
            // Count paragraphs that are part of a list.
            int listParagraphCount = 0;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                if (para.ListFormat.IsListItem)
                    listParagraphCount++;
            }

            Console.WriteLine($"Document contains {listParagraphCount} list paragraphs.");

            // Save the document to a file.
            doc.Save("NumberedList.docx");
        }
    }
}
