using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original content.");

        // Move the cursor to the end of the document's range and append new text.
        builder.MoveToDocumentEnd();
        builder.Writeln("Appended text.");

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AppendedDocument.docx");
        doc.Save(outputPath);
    }
}
