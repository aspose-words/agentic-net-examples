using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SetTableIndentation
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure paragraph indents that we want the table to match.
        builder.ParagraphFormat.LeftIndent = 72;   // 1 inch left indent.
        builder.ParagraphFormat.RightIndent = 36;  // 0.5 inch right indent.

        // Insert a table after setting the paragraph format.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Align the table's left indent with the paragraph's left indent.
        table.LeftIndent = builder.ParagraphFormat.LeftIndent;

        // Aspose.Words versions prior to 23.5 do not expose a RightIndent property.
        // To emulate a right indent we set the table's preferred width so that the
        // distance from the right edge of the page matches the paragraph's right indent.
        // PreferredWidth = PageWidth - LeftIndent - RightIndent.
        double pageWidth = builder.PageSetup.PageWidth;
        double rightIndent = builder.ParagraphFormat.RightIndent;
        table.PreferredWidth = PreferredWidth.FromPoints(pageWidth - table.LeftIndent - rightIndent);

        // Save the document.
        doc.Save("TableAlignedWithParagraphIndent.docx");
    }
}
