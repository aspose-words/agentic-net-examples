using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListModificationExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a numbered list with three items.
            List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;
            builder.Writeln("Numbered item 1");
            builder.Writeln("Numbered item 2");
            builder.Writeln("Numbered item 3");
            builder.ListFormat.RemoveNumbers();

            // Add a bulleted list with two items.
            List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);
            builder.ListFormat.List = bulletedList;
            builder.Writeln("Bullet item A");
            builder.Writeln("Bullet item B");
            builder.ListFormat.RemoveNumbers();

            // Iterate through all lists in the document and apply uniform modifications.
            foreach (List list in doc.Lists)
            {
                // Example modification: set the font color of the first level to DarkGreen.
                if (list.ListLevels.Count > 0)
                {
                    list.ListLevels[0].Font.Color = Color.DarkGreen;
                }

                // Ensure that each list restarts numbering at each new section.
                list.IsRestartAtEachSection = true;
            }

            // Save the document to the local file system.
            doc.Save("ModifiedLists.docx");
        }
    }
}
