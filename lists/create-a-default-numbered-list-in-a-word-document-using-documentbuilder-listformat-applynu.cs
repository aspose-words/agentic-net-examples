using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading for the list.
        builder.Writeln("Default numbered list example:");

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Add list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the numbered list.
        builder.ListFormat.RemoveNumbers();

        // Prepare the output folder and file name.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);
        string outputFile = Path.Combine(outputFolder, "NumberedList.docx");

        // Save the document to disk.
        doc.Save(outputFile);
    }
}
