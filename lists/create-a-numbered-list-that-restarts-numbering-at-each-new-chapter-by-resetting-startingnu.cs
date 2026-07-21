using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a numbered list based on the default template.
            List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Generate three chapters, each with its own numbered list that restarts at 1.
            for (int chapter = 1; chapter <= 3; chapter++)
            {
                // Insert a chapter heading.
                builder.ParagraphFormat.ClearFormatting();
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln($"Chapter {chapter}");

                // Reset the starting number of the first list level to 1 for this chapter.
                numberedList.ListLevels[0].StartAt = 1;

                // Apply the list to subsequent paragraphs.
                builder.ListFormat.List = numberedList;

                // Add five list items for the current chapter.
                for (int item = 1; item <= 5; item++)
                {
                    builder.Writeln($"Item {item} in Chapter {chapter}");
                }

                // End the list for this chapter.
                builder.ListFormat.RemoveNumbers();
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "NumberedList.docx");
            doc.Save(outputPath);
        }
    }
}
