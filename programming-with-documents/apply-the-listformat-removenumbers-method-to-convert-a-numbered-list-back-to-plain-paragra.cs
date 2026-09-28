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

        // Insert a simple numbered list with three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        builder.ListFormat.RemoveNumbers(); // Reset builder state.

        // Convert each numbered list item back to a regular paragraph.
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            if (para.IsListItem)
            {
                para.ListFormat.RemoveNumbers();
            }
        }

        // Save the modified document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Optional verification: reload the document and ensure no list items remain.
        Document loaded = new Document(outputPath);
        bool anyListItems = false;
        foreach (Paragraph para in loaded.FirstSection.Body.Paragraphs)
        {
            if (para.IsListItem)
            {
                anyListItems = true;
                break;
            }
        }

        // Output result status (no user interaction required).
        Console.WriteLine(anyListItems
            ? "Some paragraphs are still list items."
            : "All list numbers removed; document contains plain paragraphs.");
    }
}
