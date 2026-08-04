using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list.
            builder.ListFormat.ApplyNumberDefault();

            // Add several list items.
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");

            // End the list.
            builder.ListFormat.RemoveNumbers();

            // Define the output file path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "DefaultNumberedList.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
