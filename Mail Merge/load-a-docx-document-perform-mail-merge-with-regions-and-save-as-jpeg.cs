using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeToJpegExample
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX that contains mail‑merge regions.
            // The document must have fields like «TableStart:Employees» and «TableEnd:Employees».
            Document doc = new Document("Template.docx");

            // Prepare a DataSet with a table whose name matches the mail‑merge region.
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("FirstName");
            employeeTable.Columns.Add("LastName");
            employeeTable.Columns.Add("Title");

            // Add sample rows.
            employeeTable.Rows.Add("John", "Doe", "Sales Manager");
            employeeTable.Rows.Add("Jane", "Smith", "Marketing Director");
            employeeTable.Rows.Add("Bob", "Johnson", "Developer");

            DataSet data = new DataSet();
            data.Tables.Add(employeeTable);

            // Execute the mail merge using the DataSet. This will expand the region
            // for each row in the Employees table.
            doc.MailMerge.ExecuteWithRegions(data);

            // Save the merged document as a JPEG image.
            // ImageSaveOptions allows us to specify the format and rendering options.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            // Optional: set JPEG quality (0‑100). Higher value = better quality.
            jpegOptions.JpegQuality = 90;

            doc.Save("MergedResult.jpg", jpegOptions);
        }
    }
}
