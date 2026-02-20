using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOTM template (macros are preserved when loading a .dotm file)
        Document doc = new Document(@"C:\Templates\Template.dotm");

        // Create a DocumentBuilder attached to the loaded document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table
        Table table = builder.StartTable();

        // First row – header cells
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row – data cells
        builder.InsertCell();
        builder.Writeln("Value 1");
        builder.InsertCell();
        builder.Writeln("Value 2");
        builder.EndRow();

        // End the table
        builder.EndTable();

        // Optional: adjust table layout (auto‑fit to contents)
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document. Use SaveFormat.Docm if you need to keep the macros.
        doc.Save(@"C:\Output\Result.docx");
    }
}
