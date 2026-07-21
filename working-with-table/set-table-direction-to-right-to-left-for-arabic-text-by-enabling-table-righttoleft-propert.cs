using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a couple of cells with Arabic text.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("مرحبا"); // "Hello" in Arabic
        builder.InsertCell();
        builder.Write("عالم"); // "World" in Arabic
        builder.EndRow();

        builder.InsertCell();
        builder.Write("كيف حالك؟"); // "How are you?" in Arabic
        builder.InsertCell();
        builder.Write("بخير"); // "Fine" in Arabic
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Enable right‑to‑left layout for the table.
        table.Bidi = true;

        // Save the document to the local file system.
        doc.Save("TableRightToLeft.docx");
    }
}
