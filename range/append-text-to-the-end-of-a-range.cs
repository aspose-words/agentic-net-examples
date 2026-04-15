using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some initial text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original content.");

        // Append additional text to the end of the document's range.
        builder.MoveToDocumentEnd();
        builder.Write(" Appended text.");

        // Save the document.
        doc.Save("Result.docx");
    }
}
