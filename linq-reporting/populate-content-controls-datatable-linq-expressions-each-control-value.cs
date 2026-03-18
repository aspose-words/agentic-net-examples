using System;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlPopulation
{
    class Program
    {
        static void Main()
        {
            // Load the template document that contains content controls (StructuredDocumentTag nodes).
            // Replace "Template.docx" with the actual path to your template file.
            Document doc = new Document("Template.docx");

            // Create a sample DataTable. In a real scenario, this would come from a database,
            // CSV file, or any other source.
            DataTable table = CreateSampleDataTable();

            // Assume we want to populate the content controls with data from the first row.
            // Use LINQ to retrieve the first DataRow.
            DataRow firstRow = table.AsEnumerable().FirstOrDefault();
            if (firstRow == null)
                throw new InvalidOperationException("The DataTable does not contain any rows.");

            // Iterate over all StructuredDocumentTag nodes (content controls) in the document.
            // The Tag property is commonly used to store a key that matches a column name.
            foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
            {
                // Use the Tag as the lookup key; if Tag is empty, fall back to Title.
                string key = !string.IsNullOrEmpty(sdt.Tag) ? sdt.Tag : sdt.Title;

                // If the key is empty, we cannot map it to a column.
                if (string.IsNullOrEmpty(key))
                    continue;

                // Check if the DataTable contains a column with this key (case‑insensitive).
                DataColumn column = table.Columns
                                          .Cast<DataColumn>()
                                          .FirstOrDefault(c => string.Equals(c.ColumnName, key, StringComparison.OrdinalIgnoreCase));

                if (column == null)
                    continue; // No matching column; skip this content control.

                // Retrieve the value from the DataRow for the matched column.
                object value = firstRow[column];

                // Convert the value to string (handle nulls gracefully).
                string text = value?.ToString() ?? string.Empty;

                // Populate the content control.
                // For plain‑text or rich‑text controls we can set the Text directly.
                // For other types (e.g., dropdown, checkbox) additional handling would be required.
                sdt.RemoveAllChildren(); // Clear existing content.
                Paragraph para = new Paragraph(doc);
                para.AppendChild(new Run(doc, text));
                sdt.AppendChild(para);
            }

            // Save the populated document.
            // Replace "Result.docx" with the desired output path.
            doc.Save("Result.docx");
        }

        /// <summary>
        /// Generates a sample DataTable with columns that correspond to content control tags.
        /// </summary>
        private static DataTable CreateSampleDataTable()
        {
            DataTable dt = new DataTable("SampleData");
            dt.Columns.Add("FirstName");
            dt.Columns.Add("LastName");
            dt.Columns.Add("Email");

            // Add a single row of data.
            dt.Rows.Add("John", "Doe", "john.doe@example.com");

            return dt;
        }
    }
}
