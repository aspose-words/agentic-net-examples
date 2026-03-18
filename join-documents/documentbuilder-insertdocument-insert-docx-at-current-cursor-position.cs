using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the files used in the example.
        const string mainDocPath = "Main.docx";          // The base document (can be empty or pre‑existing).
        const string insertDocPath = "Insert.docx";      // The DOCX file that will be inserted.
        const string outputHtmlPath = "Result.html";    // The final HTML file.

        // -----------------------------------------------------------------
        // 1. Create (or load) the destination document.
        // -----------------------------------------------------------------
        // Here we create a new blank document using the Document() constructor.
        Document destination = new Document();

        // Use DocumentBuilder to add some initial content and to obtain a cursor.
        DocumentBuilder builder = new DocumentBuilder(destination);
        builder.Writeln("Content before the inserted document.");

        // -----------------------------------------------------------------
        // 2. Load the source document that will be inserted.
        // -----------------------------------------------------------------
        // The Document(string) constructor loads an existing DOCX from disk.
        Document source = new Document(insertDocPath);

        // -----------------------------------------------------------------
        // 3. Insert the source document at the current cursor position.
        // -----------------------------------------------------------------
        // InsertDocument(Document, ImportFormatMode) inserts the whole document
        // and keeps the original formatting of the source.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Add some trailing content to demonstrate that insertion happened at the cursor.
        builder.Writeln("Content after the inserted document.");

        // -----------------------------------------------------------------
        // 4. Save the combined document as HTML.
        // -----------------------------------------------------------------
        // Save(string, SaveFormat) determines the output format from the enum.
        destination.Save(outputHtmlPath, SaveFormat.Html);
    }
}
