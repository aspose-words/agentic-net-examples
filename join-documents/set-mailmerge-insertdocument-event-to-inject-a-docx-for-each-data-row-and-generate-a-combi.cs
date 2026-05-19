using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create two sample source documents that will be inserted.
        string docPath1 = Path.Combine(workDir, "Source1.docx");
        string docPath2 = Path.Combine(workDir, "Source2.docx");
        CreateSampleDocument(docPath1, "Content of the first inserted document.");
        CreateSampleDocument(docPath2, "Content of the second inserted document.");

        // List of documents to insert – order corresponds to data rows.
        List<string> documentsToInsert = new List<string> { docPath1, docPath2 };

        // Create the mail‑merge template containing a document merge field.
        string templatePath = Path.Combine(workDir, "Template.docx");
        CreateTemplateDocument(templatePath);

        // Load the template.
        Document template = new Document(templatePath);

        // Attach a field‑merging callback that inserts a document for each record.
        template.MailMerge.FieldMergingCallback = new InsertDocumentCallback(documentsToInsert);

        // Prepare a simple data source – the actual values are not used,
        // they only drive the number of records.
        DataTable data = new DataTable("Data");
        data.Columns.Add("Id", typeof(int));
        data.Rows.Add(1);
        data.Rows.Add(2);

        // Execute the mail merge; each row will cause the callback to fire.
        template.MailMerge.Execute(data);

        // Save the combined result as PDF.
        string outputPdf = Path.Combine(workDir, "Combined.pdf");
        template.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The combined PDF was not created.");
    }

    // Creates a one‑page DOCX with the specified text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);
    }

    // Creates a template DOCX that contains a document merge field.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The \b switch marks the field as a document merge field.
        builder.InsertField("MERGEFIELD Document \\b");
        builder.Writeln(); // Ensure the field is on its own line.
        doc.Save(filePath, SaveFormat.Docx);
    }

    // Callback that inserts a document at the merge field location.
    private class InsertDocumentCallback : IFieldMergingCallback
    {
        private readonly List<string> _documents;

        public InsertDocumentCallback(List<string> documents)
        {
            _documents = documents ?? throw new ArgumentNullException(nameof(documents));
        }

        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Only handle the specific document merge field.
            if (!args.DocumentFieldName.Equals("Document", StringComparison.OrdinalIgnoreCase))
                return;

            int recordIndex = args.RecordIndex; // zero‑based index of the current data row.
            if (recordIndex < 0 || recordIndex >= _documents.Count)
                return;

            // Load the document to be inserted.
            Document insertDoc = new Document(_documents[recordIndex]);

            // Move the builder to the merge field and insert the document.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.DocumentFieldName);
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // Suppress the default text insertion for this field.
            args.Text = string.Empty;
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
