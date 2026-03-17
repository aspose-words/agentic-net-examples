using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a styled paragraph (Heading1) that will precede the table.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Table will be inserted below:");

        // Insert an empty paragraph that will become the container for the table.
        builder.Writeln(); // This creates a new paragraph and moves the cursor into it.

        // Move the builder back to the newly created paragraph.
        // Paragraph indices are zero‑based: section index 0, paragraph index = last paragraph.
        int paragraphIndex = doc.FirstSection.Body.Paragraphs.Count - 1;
        builder.MoveToParagraph(0, paragraphIndex);

        // Start building the table inside the current paragraph.
        Table table = builder.StartTable();

        // ---- First row (header) ----
        builder.InsertCell();               // First cell of the row.
        builder.Write("Header 1");          // Write content into the cell.
        builder.InsertCell();               // Second cell of the row.
        builder.Write("Header 2");
        builder.EndRow();                   // End the header row.

        // ---- Second row (data) ----
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();                   // End the data row.

        // Finish the table construction.
        builder.EndTable();

        // Save the resulting document.
        doc.Save("TableInStyledParagraph.docx");
    }
}
