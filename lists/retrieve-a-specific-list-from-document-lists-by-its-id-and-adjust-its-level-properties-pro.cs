using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListRetrievalExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a bullet list using a predefined template.
            List list = doc.Lists.Add(ListTemplate.BulletDefault);

            // Apply the list to the builder.
            builder.ListFormat.List = list;

            // Write some list items.
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Retrieve the same list by its unique ListId.
            int listId = list.ListId;
            List retrievedList = doc.Lists.GetListByListId(listId);

            // Adjust level properties of the first level (index 0).
            ListLevel level = retrievedList.ListLevels[0];
            level.Font.Name = "Arial";
            level.Font.Size = 12;
            level.Alignment = ListLevelAlignment.Right; // Align the bullet/number to the right.
            level.NumberStyle = NumberStyle.Bullet; // Ensure bullet style remains.
            level.NumberPosition = 20; // Position of the bullet.
            level.TextPosition = 40;   // Position of the text after the bullet.

            // Save the document to a file.
            doc.Save("AdjustedList.docx");

            // Indicate completion.
            Console.WriteLine("Document created and list level adjusted successfully.");
        }
    }
}
