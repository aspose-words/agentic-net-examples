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

        // Set paragraph indentation (in points).
        builder.ParagraphFormat.LeftIndent = 30;   // Left indent for the paragraph.
        builder.ParagraphFormat.RightIndent = 20;  // Right indent for the paragraph.

        // Add a paragraph to visualize the indentation.
        builder.Writeln("This paragraph demonstrates left and right indents.");

        // Start a table.
        Table table = builder.StartTable();

        // Insert the first cell to create the initial row (required before setting table formatting).
        builder.InsertCell();

        // Align the table's left indent with the paragraph's left indent.
        table.LeftIndent = builder.ParagraphFormat.LeftIndent;

        // Insert text into the first cell.
        builder.Writeln("Table cell aligned with paragraph left indent.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the local file system.
        doc.Save("TableIndentExample.docx");
    }
}
