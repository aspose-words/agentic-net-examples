using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace RemoveListNumbersExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.Writeln("Item 3");

            // Stop list formatting for the current paragraph and subsequent ones.
            builder.ListFormat.RemoveNumbers();

            // This paragraph will not have numbering.
            builder.Writeln("This paragraph is not part of the list.");

            // Save the document to the current directory.
            string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "RemoveNumbers.docx");
            doc.Save(outputPath);
        }
    }
}
