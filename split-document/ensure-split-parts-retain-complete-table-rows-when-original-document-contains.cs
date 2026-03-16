using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Ensure that no table row is allowed to break across a page.
        // This keeps each row intact when the document is split.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            foreach (Row row in table.Rows)
                row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Configure the save options to split the document at page breaks.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak
        };

        // Save the document; Aspose.Words will create separate HTML parts,
        // each containing whole table rows.
        doc.Save("Output.html", saveOptions);
    }
}
