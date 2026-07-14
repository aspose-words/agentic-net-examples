using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several paragraphs inside a loop using Writeln.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is paragraph number {i}.");
        }

        // Save the resulting document.
        doc.Save("InsertedParagraphs.docx");
    }
}
