using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListLevelExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a numbered list using the default list template.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

            // Set the list level to 3 (zero‑based levels: 0 = first level, 1 = second, 2 = third, 3 = fourth).
            // According to the task description we set it to three.
            builder.ListFormat.ListLevelNumber = 3;

            // Add a paragraph that will appear as a fourth‑level item in the list.
            builder.Writeln("This is a list item at level 3.");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Define output path (in the same folder as the executable).
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ListLevelExample.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
