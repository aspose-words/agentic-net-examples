using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a numbered list.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Numbered item 1");
            builder.Writeln("Numbered item 2");
            builder.ListFormat.RemoveNumbers();

            // Add a bulleted list.
            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("Bulleted item 1");
            builder.Writeln("Bulleted item 2");
            builder.ListFormat.RemoveNumbers();

            // Apply a uniform style to all lists in the document.
            foreach (List list in doc.Lists)
            {
                foreach (ListLevel level in list.ListLevels)
                {
                    level.Font.Name = "Arial";
                    level.Font.Color = Color.DarkGreen;
                    level.Font.Bold = true;
                }
            }

            // Save the document.
            doc.Save("UniformListStyle.docx");
        }
    }
}
