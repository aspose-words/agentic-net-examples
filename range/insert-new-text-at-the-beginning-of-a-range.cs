using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths for the sample source and the final output.
        string sourcePath = Path.Combine(Environment.CurrentDirectory, "Source.docx");
        string resultPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with some initial content.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Original content of the document.");

        // Save the source document locally (bootstrap step).
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the previously saved document.
        // -----------------------------------------------------------------
        Document doc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Insert new text at the beginning of the document's range.
        //    Moving the builder's cursor to the start of the document places it
        //    before the first character of the document range.
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(doc);
        insertBuilder.MoveToDocumentStart();
        insertBuilder.Write("Inserted at start. ");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(resultPath);

        // Optional verification (not required for the task, but harmless).
        // The following line demonstrates that the text was inserted.
        string finalText = doc.Range.Text.Trim();
        Console.WriteLine("Final document text: " + finalText);
    }
}
