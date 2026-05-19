using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsBulletedList
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a bulleted list based on the default bullet template.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

            // Customize the first level bullet to use a dash ("-").
            ListLevel level = bulletList.ListLevels[0];
            level.NumberStyle = NumberStyle.Bullet;   // Ensure the level is a bullet.
            level.NumberFormat = "-";                 // Set the bullet character.

            // Apply the customized list to the builder.
            builder.ListFormat.List = bulletList;

            // Add three list items.
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to a file in the current directory.
            doc.Save("BulletedList.docx");
        }
    }
}
