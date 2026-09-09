using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Use DocumentBuilder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list.
            builder.ListFormat.ApplyNumberDefault();

            // Add a few list items – these paragraphs will be numbered.
            builder.Writeln("Numbered item 1");
            builder.Writeln("Numbered item 2");
            builder.Writeln("Numbered item 3");

            // Stop list formatting for subsequent paragraphs.
            // This call removes numbers/bullets from the current paragraph and resets the list level.
            builder.ListFormat.RemoveNumbers();

            // Add a normal paragraph that is not part of the list.
            builder.Writeln("This paragraph is not numbered.");

            // Save the document to a file in the current directory.
            doc.Save("Lists.RemoveNumbers.docx");
        }
    }
}
