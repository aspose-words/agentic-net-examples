using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace CustomBulletListExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a new list based on the bullet template.
            List list = doc.Lists.Add(ListTemplate.BulletDefault);

            // Configure the first level of the list to use a custom bullet character.
            ListLevel level = list.ListLevels[0];
            // The bullet template already uses the Bullet numbering style,
            // so we only need to change the character that is displayed.
            level.NumberFormat = "★"; // Custom bullet character.

            // Apply the list to the builder.
            builder.ListFormat.List = list;

            // Write list items.
            builder.Writeln("First custom bullet item");
            builder.Writeln("Second custom bullet item");
            builder.Writeln("Third custom bullet item");

            // Remove list formatting after the list is complete.
            builder.ListFormat.RemoveNumbers();

            // Add a normal paragraph after the list to verify list formatting is not applied unintentionally.
            builder.Writeln("This paragraph is not part of the list.");

            // Save the document to a file.
            string outputPath = "CustomBulletList.docx";
            doc.Save(outputPath);

            // Validation: count paragraphs that are part of the list.
            int listParagraphCount = 0;
            foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
            {
                if (para.ListFormat.IsListItem)
                    listParagraphCount++;
            }

            // Output validation results.
            Console.WriteLine($"Document saved to '{outputPath}'.");
            Console.WriteLine($"Number of list paragraphs: {listParagraphCount}");
            Console.WriteLine("Expected number of list paragraphs: 3");
        }
    }
}
