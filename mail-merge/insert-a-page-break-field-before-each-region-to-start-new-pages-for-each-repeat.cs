using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsMailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a data table that will be used as the mail merge data source.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FullName");
            table.Columns.Add("Address");
            table.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London" });
            table.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a page break before the mail merge region so each repeat starts on a new page.
            builder.InsertBreak(BreakType.PageBreak);

            // Insert the start of the mail merge region.
            builder.InsertField($" MERGEFIELD TableStart:{table.TableName}");

            // Insert the fields that will be populated from the data source.
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Write(", ");
            builder.InsertField(" MERGEFIELD Address ");

            // Insert the end of the mail merge region.
            builder.InsertField($" MERGEFIELD TableEnd:{table.TableName}");

            // Execute the mail merge with regions. The region will be repeated for each row,
            // each on a new page because of the page break inserted before the region.
            doc.MailMerge.ExecuteWithRegions(table);

            // Save the resulting document.
            doc.Save("MailMergeWithPageBreak.docx");
        }
    }
}
