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

            // Create a numbered list based on the default list template.
            List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Configure the list to restart numbering at each new section.
            // Each chapter will be placed in its own section, so the list restarts automatically.
            numberedList.IsRestartAtEachSection = true;

            // Define how many chapters and items per chapter we want.
            int chapterCount = 3;
            int itemsPerChapter = 4;

            for (int chapter = 1; chapter <= chapterCount; chapter++)
            {
                // Insert a section break before every chapter except the first one.
                if (chapter > 1)
                {
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
                }

                // Write the chapter heading using the built‑in Heading 1 style.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln($"Chapter {chapter}");

                // Switch back to the normal style for the list items.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

                // Apply the previously created list to the following paragraphs.
                builder.ListFormat.List = numberedList;

                // Add the list items for this chapter.
                for (int item = 1; item <= itemsPerChapter; item++)
                {
                    builder.Writeln($"Item {item} of Chapter {chapter}");
                }

                // End the list for this chapter.
                builder.ListFormat.RemoveNumbers();
            }

            // Save the document to the output folder.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "NumberedListRestartByChapter.docx");
            doc.Save(outputPath);
        }
    }
}
