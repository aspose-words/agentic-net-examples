using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeToPngDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file that contains MERGEFIELDs.
            const string sourceDocPath = "Template.docx";

            // Load the source document.
            Document sourceDoc = new Document(sourceDocPath);

            // -----------------------------------------------------------------
            // 1. Mail merge using an array of field names and values (single record).
            // -----------------------------------------------------------------
            Document docArray = (Document)sourceDoc.Clone();

            // Example field names – they must exist in the template.
            string[] fieldNames = { "FullName", "Company" };
            object[] fieldValues = { "James Bond", "MI5 Headquarters" };

            // Perform the mail merge.
            docArray.MailMerge.Execute(fieldNames, fieldValues);

            // Save the result as a PNG image (first page only).
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the first page.
                PageSet = new PageSet(0)
            };
            docArray.Save("MailMerge_Array.png", pngOptions);

            // -----------------------------------------------------------------
            // 2. Mail merge using a DataTable (multiple records).
            // -----------------------------------------------------------------
            Document docTable = (Document)sourceDoc.Clone();

            // Build a DataTable with the same columns as the merge fields.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FullName");
            table.Columns.Add("Company");

            // Add sample rows.
            table.Rows.Add("John Doe", "Acme Corp");
            table.Rows.Add("Jane Smith", "Globex Inc");

            // Perform the mail merge – this will create a separate document section for each row.
            docTable.MailMerge.Execute(table);

            // Save the merged document as PNG (each page will be saved separately).
            // Here we render all pages one by one.
            for (int i = 0; i < docTable.PageCount; i++)
            {
                pngOptions.PageSet = new PageSet(i);
                string outFile = $"MailMerge_Table_Page{i + 1}.png";
                docTable.Save(outFile, pngOptions);
            }

            // -----------------------------------------------------------------
            // 3. Mail merge using a single DataRow (one record).
            // -----------------------------------------------------------------
            Document docRow = (Document)sourceDoc.Clone();

            // Reuse the first row from the DataTable created above.
            DataRow row = table.Rows[0];

            // Perform the mail merge with a single DataRow.
            docRow.MailMerge.Execute(row);

            // Save the result as PNG.
            pngOptions.PageSet = new PageSet(0);
            docRow.Save("MailMerge_DataRow.png", pngOptions);
        }
    }
}
