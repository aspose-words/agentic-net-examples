using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeInsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Create an empty document that will receive the inserted documents.
            Document mainDoc = new Document();

            // Prepare a data source – each row will cause one document to be inserted.
            DataTable data = new DataTable("Docs");
            data.Columns.Add("FileName"); // Name of the DOCX file (without extension) to insert.

            // Example rows – adjust the file names and paths as needed.
            data.Rows.Add("Contract_A");
            data.Rows.Add("Contract_B");
            data.Rows.Add("Contract_C");

            // Iterate through each data row, load the corresponding document and append it to the main document.
            foreach (DataRow row in data.Rows)
            {
                string fileName = row["FileName"].ToString();
                string sourcePath = $"SourceDocs\\{fileName}.docx";

                // Load the source document.
                Document srcDoc = new Document(sourcePath);

                // Append the source document to the main document.
                // KeepSourceFormatting preserves the original layout of each inserted document.
                mainDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            }

            // Save the combined result as a PDF file.
            mainDoc.Save("CombinedResult.pdf", SaveFormat.Pdf);
        }
    }
}
