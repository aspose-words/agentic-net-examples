using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build initial content so we have something to insert before.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original content of the document.");

        // Insert new text at the very beginning of the document's range.
        DocumentBuilder inserter = new DocumentBuilder(doc);
        inserter.MoveToDocumentStart();               // Position the cursor at the start.
        inserter.Write("Inserted at start. ");        // Write the new text.

        // Optional verification: the document text should start with the inserted string.
        string fullText = doc.GetText();
        if (!fullText.StartsWith("Inserted at start.", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Text was not inserted at the beginning of the range.");
        }

        // Save the modified document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "InsertedText.docx");
        doc.Save(outputPath);
    }
}
