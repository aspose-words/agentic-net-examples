using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsBulletedListExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a new single‑level bulleted list based on the default bullet template.
            List bulletList = doc.Lists.AddSingleLevelList(ListTemplate.BulletDefault);

            // Customize the first (and only) level of the list to use a dash ("-") as the bullet character.
            // NumberStyle.Bullet tells Word that this is a bullet list.
            bulletList.ListLevels[0].NumberStyle = NumberStyle.Bullet;
            bulletList.ListLevels[0].NumberFormat = "-";

            // Apply the custom list to the builder.
            builder.ListFormat.List = bulletList;

            // Add three list items.
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Determine an output path in the current working directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BulletedList.docx");

            // Save the document to the specified file.
            doc.Save(outputPath);
        }
    }
}
