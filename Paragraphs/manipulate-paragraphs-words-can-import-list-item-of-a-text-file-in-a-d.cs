using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists; // Needed for ListTemplate

namespace AsposeWordsListImport
{
    class Program
    {
        static void Main()
        {
            // Path to the plain‑text file that contains one list item per line.
            string textFilePath = @"C:\Data\list_items.txt";

            // Path where the resulting DOCX document will be saved.
            string outputDocPath = @"C:\Data\ListDocument.docx";

            // Load all lines from the text file.
            // Each line will become a separate list item in the Word document.
            string[] lines = File.ReadAllLines(textFilePath);

            // Create a new blank Word document.
            Document doc = new Document();

            // Create a DocumentBuilder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a heading for the list (optional).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Imported List Items");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Create a new bullet list in the document.
            // The list is added to the document's Lists collection and assigned to the builder.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);

            // Insert each line as a separate list item.
            foreach (string line in lines)
            {
                // Write the line and automatically create a new paragraph (list item).
                builder.Writeln(line);
            }

            // Stop list formatting for any further content.
            builder.ListFormat.List = null;               // Detach from the list
            builder.ListFormat.ListLevelNumber = 0;       // Reset the level

            // Save the document to the specified file.
            doc.Save(outputDocPath);
        }
    }
}
