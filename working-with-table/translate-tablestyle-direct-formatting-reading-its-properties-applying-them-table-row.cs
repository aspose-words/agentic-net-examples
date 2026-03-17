using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleToDirectFormattingDemo
{
    class Program
    {
        static void Main()
        {
            // Load a DOCX document that contains tables formatted with table styles.
            Document doc = new Document("TablesWithStyles.docx");

            // Convert all formatting defined in the table styles into direct formatting
            // on the tables, rows, and cells. This expands the style so that the
            // formatting can be queried or modified without relying on the style.
            doc.ExpandTableStylesToDirectFormatting();

            // Save the document. The tables now have their formatting applied directly.
            doc.Save("TablesDirectFormatting.docx");
        }
    }
}
