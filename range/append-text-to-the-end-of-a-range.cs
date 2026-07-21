using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text.
        builder.Writeln("Hello world!");

        // Move the cursor to the end of the document and append additional text.
        builder.MoveToDocumentEnd();
        builder.Write(" Appended text.");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
