using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListTemplateExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a heading before the lists.
            builder.Writeln("Document with custom list template (max 9 levels)");

            // Define the custom list template to be applied to all lists.
            // For this example we use NumberUppercaseLetterDot (A., B., C., ...).
            ListTemplate customTemplate = ListTemplate.NumberUppercaseLetterDot;

            // Create three separate lists to demonstrate the template application.
            for (int listIndex = 1; listIndex <= 3; listIndex++)
            {
                // Add a new list based on the custom template.
                List list = doc.Lists.Add(customTemplate);

                // Apply the list to the builder.
                builder.ListFormat.List = list;

                // Write list items with nesting levels.
                // Enforce a maximum of nine nesting levels (0‑8).
                for (int level = 0; level < 12; level++) // Attempt more than 9 levels.
                {
                    if (level >= 9) // Stop adding items beyond the ninth level.
                        break;

                    builder.ListFormat.ListLevelNumber = level;
                    builder.Writeln($"List {listIndex}, Level {level + 1}");
                }

                // End the current list.
                builder.ListFormat.RemoveNumbers();

                // Add a blank line between lists.
                builder.Writeln();
            }

            // Save the document to the local file system.
            string outputPath = "CustomListTemplate.docx";
            doc.Save(outputPath);
        }
    }
}
