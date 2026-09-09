using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListIndentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list.
            builder.ListFormat.ApplyNumberDefault();

            // Write several paragraphs, increasing the list level before each new paragraph.
            // The first paragraph is at level 0, then we indent for the next one, and so on.
            for (int i = 0; i < 5; i++)
            {
                builder.Writeln($"Item at list level {i}");

                // Increase the list level for the next paragraph, except after the last one.
                if (i < 4)
                {
                    builder.ListFormat.ListIndent();
                }
            }

            // Reset the list level back to the base level and end the list.
            while (builder.ListFormat.ListLevelNumber > 0)
            {
                builder.ListFormat.ListOutdent();
            }
            builder.ListFormat.RemoveNumbers();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "IncreaseListIndent.docx");
            doc.Save(outputPath);
        }
    }
}
