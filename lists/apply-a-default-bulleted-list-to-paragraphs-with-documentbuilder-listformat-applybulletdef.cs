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

            // Attach a DocumentBuilder to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a heading before the list.
            builder.Writeln("Default bulleted list example:");

            // Start a default bulleted list and apply it to subsequent paragraphs.
            builder.ListFormat.ApplyBulletDefault();

            // Add several list items.
            builder.Writeln("First bullet item");
            builder.Writeln("Second bullet item");
            builder.Writeln("Third bullet item");
            builder.Writeln("Fourth bullet item");
            builder.Writeln("Fifth bullet item");

            // End the list.
            builder.ListFormat.RemoveNumbers();

            // Define the output folder and file name.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "DefaultBulletedList.docx");

            // Save the document to the specified path.
            doc.Save(outputPath);
        }
    }
}
