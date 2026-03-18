using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document that contains the table.
        Document doc = new Document("Table spanning two pages.docx");

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Prevent each row from breaking across a page break.
        // This keeps the rows together, which effectively keeps the whole table on one page
        // as long as the table fits within the page height.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the modified document.
        doc.Save("Table.KeepTogether.docx");
    }
}
