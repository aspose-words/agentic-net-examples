using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;
using System.Drawing;

class RestartNumberingByChapter
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on the default template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Define the number of chapters.
        int chapterCount = 3;

        for (int chapter = 1; chapter <= chapterCount; chapter++)
        {
            // Insert a chapter heading (e.g., "Chapter 1").
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {chapter}");

            // Reset the starting number for the first level of the list.
            // This makes the list start from 1 for each new chapter.
            numberedList.ListLevels[0].StartAt = 1;

            // Apply the list to subsequent paragraphs.
            builder.ListFormat.List = numberedList;

            // Add some list items for the current chapter.
            for (int i = 1; i <= 5; i++)
            {
                builder.Writeln($"Item {i} of Chapter {chapter}");
            }

            // Remove list formatting before the next heading.
            builder.ListFormat.RemoveNumbers();
        }

        // Save the document.
        doc.Save("RestartNumberingByChapter.docx");
    }
}
