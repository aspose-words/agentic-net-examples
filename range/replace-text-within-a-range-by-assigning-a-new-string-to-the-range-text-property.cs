using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Hello World! This is a sample text.");

        // Replace the word "sample" with "example" using the document's range.
        doc.Range.Replace("sample", "example");

        // Save the modified document to the local file system.
        doc.Save("Output.docx");
    }
}
