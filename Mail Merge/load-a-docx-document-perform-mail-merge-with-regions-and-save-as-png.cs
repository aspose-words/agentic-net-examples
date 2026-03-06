using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeToPngExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file that contains mail‑merge regions.
            const string inputPath = @"C:\Docs\TemplateWithRegions.docx";

            // Load the existing document. This uses the provided Document(string) constructor.
            Document doc = new Document(inputPath);

            // Prepare a DataSet whose table name matches the mail‑merge region name in the document.
            // For example, the document contains fields:
            //   MERGEFIELD TableStart:Employees
            //   MERGEFIELD FirstName
            //   MERGEFIELD LastName
            //   MERGEFIELD TableEnd:Employees
            DataSet data = CreateSampleDataSet();

            // Execute the mail merge with regions. This uses the MailMerge.ExecuteWithRegions(DataSet) method.
            doc.MailMerge.ExecuteWithRegions(data);

            // Configure image save options to render each page as a PNG image.
            // The ImageSaveOptions class is part of the provided Save overloads.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render all pages; each page will be saved as a separate PNG file.
                // The {0} placeholder in the file name will be replaced with the page number (1‑based).
                // If you want a single image of the first page only, set PageSet = new PageSet(0);
                // Here we keep the default to render all pages.
                // You can also adjust resolution, size, etc., if required.
                Resolution = 300
            };

            // Save the document as PNG images. The file name pattern includes a page number placeholder.
            // This uses the Document.Save(string, SaveOptions) overload.
            const string outputPattern = @"C:\Docs\ResultPage_{0}.png";
            doc.Save(outputPattern, pngOptions);

            // If you prefer to save only the first page as a single PNG, uncomment the following:
            // pngOptions.PageSet = new PageSet(0);
            // doc.Save(@"C:\Docs\ResultFirstPage.png", pngOptions);
        }

        /// <summary>
        /// Creates a sample DataSet with a single DataTable named "Employees".
        /// The table columns correspond to the merge fields inside the region.
        /// </summary>
        private static DataSet CreateSampleDataSet()
        {
            // Define the table that matches the region name.
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("FirstName", typeof(string));
            employees.Columns.Add("LastName", typeof(string));

            // Add sample rows.
            employees.Rows.Add("John", "Doe");
            employees.Rows.Add("Jane", "Smith");
            employees.Rows.Add("Bob", "Johnson");

            // Assemble the DataSet.
            DataSet ds = new DataSet();
            ds.Tables.Add(employees);

            return ds;
        }
    }
}
