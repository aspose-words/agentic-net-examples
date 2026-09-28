using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three paragraphs that will become bulleted list items.
        for (int i = 1; i <= 3; i++)
        {
            // Start a new paragraph.
            builder.Writeln($"List item {i}");

            // Apply the default bullet formatting to the current paragraph.
            builder.ListFormat.ApplyBulletDefault();
        }

        // Save the document to the output file.
        doc.Save("BulletedList.docx");
    }
}
