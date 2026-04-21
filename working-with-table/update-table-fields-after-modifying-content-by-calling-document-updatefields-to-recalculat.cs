using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsTableFieldUpdate
{
    public class Program
    {
        public static void Main()
        {
            // Define output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(outputDir);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set an initial value for a built‑in document property.
            // The field we will insert later will display this value.
            doc.BuiltInDocumentProperties.Category = "Initial Category";

            // Build a simple 1‑row, 2‑column table.
            // First cell: label, second cell: a DOCPROPERTY field that shows the Category property.
            builder.StartTable();

            // Cell 1 – label.
            builder.InsertCell();
            builder.Write("Category:");

            // Cell 2 – field (do not update immediately).
            builder.InsertCell();
            // Insert the field without updating it so that the result is empty for now.
            // Use the overload that accepts a field code and a placeholder value.
            builder.InsertField(" DOCPROPERTY Category ", "");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document before updating fields (optional, just for demonstration).
            string beforePath = Path.Combine(outputDir, "TableWithField_BeforeUpdate.docx");
            doc.Save(beforePath);

            // Modify the underlying data – change the document property value.
            doc.BuiltInDocumentProperties.Category = "Updated Category";

            // Recalculate all fields in the document, including the one inside the table.
            doc.UpdateFields();

            // Save the updated document.
            string afterPath = Path.Combine(outputDir, "TableWithField_AfterUpdate.docx");
            doc.Save(afterPath);
        }
    }
}
