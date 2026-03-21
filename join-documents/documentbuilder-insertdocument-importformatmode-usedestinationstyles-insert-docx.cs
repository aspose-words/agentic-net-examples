using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace InsertDocAndSaveHtml
{
    class Program
    {
        static void Main()
        {
            // Create the styled template document in memory.
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            // Apply a heading style to demonstrate destination styles.
            templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            templateBuilder.Writeln("Template Heading");
            // Add a normal paragraph.
            templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            templateBuilder.Writeln("Template body text.");

            // Create the source document that will be inserted.
            Document sourceDoc = new Document();
            DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
            sourceBuilder.Writeln("Inserted paragraph 1.");
            sourceBuilder.Writeln("Inserted paragraph 2 with a different style.");
            // Apply a different style to show style import.
            sourceBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
            sourceBuilder.Writeln("Inserted quoted text.");

            // Insert the source document into the template using UseDestinationStyles.
            DocumentBuilder insertBuilder = new DocumentBuilder(templateDoc);
            insertBuilder.MoveToDocumentEnd();
            insertBuilder.InsertDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);

            // Save the combined document as HTML in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.html");
            templateDoc.Save(outputPath, SaveFormat.Html);

            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
