using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    // Paths for temporary files.
    private const string TemplatePath = "Template.docx";
    private const string InsertDocPath = "InsertDoc.docx";
    private const string OutputPdfPath = "Combined.pdf";

    public static void Main()
    {
        // Create sample source documents.
        CreateTemplateDocument();
        CreateInsertDocument();

        // Load the document that will be inserted for each record.
        Document insertDoc = new Document(InsertDocPath);

        // Prepare a data source with two rows (the actual field value is not used).
        DataTable data = new DataTable("Data");
        data.Columns.Add("Doc"); // Column name matches the MERGEFIELD in the template.
        data.Rows.Add(string.Empty);
        data.Rows.Add(string.Empty);

        Document? resultDoc = null;
        bool first = true;

        // Process each data row individually.
        foreach (DataRow row in data.Rows)
        {
            // Load a fresh copy of the template for the current record.
            Document tempDoc = new Document(TemplatePath);

            // Perform mail merge for the current row.
            tempDoc.MailMerge.Execute(row);

            // Replace the MERGEFIELD "Doc" with the content of the insert document.
            DocumentBuilder builder = new DocumentBuilder(tempDoc);
            builder.MoveToMergeField("Doc");
            // Insert the whole document at the merge field location.
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // Combine the processed document into the final result.
            if (first)
            {
                resultDoc = tempDoc;
                first = false;
            }
            else
            {
                // Append while preserving source formatting.
                resultDoc!.AppendDocument(tempDoc, ImportFormatMode.KeepSourceFormatting);
            }
        }

        // Save the combined document as PDF.
        if (resultDoc == null)
            throw new InvalidOperationException("No documents were merged.");

        resultDoc.Save(OutputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(OutputPdfPath))
            throw new InvalidOperationException($"Failed to create the output file '{OutputPdfPath}'.");
    }

    // Creates a simple template with a MERGEFIELD named "Doc".
    private static void CreateTemplateDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Start of the combined document.");
        // The MERGEFIELD triggers the insertion point for the additional document.
        builder.InsertField("MERGEFIELD Doc");
        builder.Writeln("End of the combined document.");

        doc.Save(TemplatePath, SaveFormat.Docx);
    }

    // Creates a simple document that will be inserted for each data row.
    private static void CreateInsertDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("=== Inserted Document Content ===");
        builder.Writeln($"Inserted at {DateTime.Now}");

        doc.Save(InsertDocPath, SaveFormat.Docx);
    }
}
