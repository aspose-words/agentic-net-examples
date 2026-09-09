using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content.
        builder.Writeln("Original text.");

        // Move the cursor to the end of the document and append additional text.
        builder.MoveToDocumentEnd();
        builder.Write(" Appended text.");

        // Save the resulting document.
        doc.Save("AppendText.docx");

        // Output the full text of the document to verify the append operation.
        Console.WriteLine(doc.Range.Text.Trim());
    }
}
