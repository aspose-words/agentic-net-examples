using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

public class Program
{
    // Paths for generated files.
    private const string TemplatePath = "Template.docx";
    private const string InsertDocPath = "InsertDoc.docx";
    private const string OutputPdfPath = "MergedOutput.pdf";

    public static void Main()
    {
        // Create a simple template with two merge fields.
        Document template = CreateTemplateDocument();
        template.Save(TemplatePath, SaveFormat.Docx);

        // Create the document that will be inserted during the mail merge.
        Document insertDoc = CreateInsertDocument();
        insertDoc.Save(InsertDocPath, SaveFormat.Docx);

        // Prepare data for the mail merge.
        DataTable data = new DataTable("Records");
        data.Columns.Add("RecordName", typeof(string));
        data.Columns.Add("InsertDoc", typeof(string)); // Value not used; callback handles insertion.
        data.Rows.Add("First Record", string.Empty);
        data.Rows.Add("Second Record", string.Empty);

        // Set up the field merging callback to insert the external DOCX.
        template.MailMerge.FieldMergingCallback = new InsertDocCallback(InsertDocPath);

        // Execute the mail merge.
        template.MailMerge.Execute(data);

        // Save the merged document as PDF.
        template.Save(OutputPdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(OutputPdfPath))
            throw new InvalidOperationException("The PDF output was not created.");

        // Optional validation: check that the merged document contains the inserted text.
        string mergedText = template.GetText();
        if (!mergedText.Contains("Content from inserted document."))
            throw new InvalidOperationException("The inserted document content was not found in the merged result.");
    }

    // Creates the mail merge template with two fields.
    private static Document CreateTemplateDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Record:");
        builder.InsertField("MERGEFIELD RecordName", "<RecordName>");
        builder.Writeln();
        builder.InsertField("MERGEFIELD InsertDoc", string.Empty);
        builder.Writeln();

        return doc;
    }

    // Creates the document that will be inserted during the merge.
    private static Document CreateInsertDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set language to en‑US using LocaleId (LCID = 1033).
        builder.Font.LocaleId = CultureInfo.GetCultureInfo("en-US").LCID;
        builder.Writeln("Content from inserted document.");

        return doc;
    }

    // Callback that inserts the external DOCX at the location of the InsertDoc merge field.
    private class InsertDocCallback : IFieldMergingCallback
    {
        private readonly string _insertDocPath;

        public InsertDocCallback(string insertDocPath)
        {
            _insertDocPath = insertDocPath;
        }

        public void FieldMerging(FieldMergingArgs args)
        {
            // Handle only the specific merge field.
            if (!args.DocumentFieldName.Equals("InsertDoc", StringComparison.OrdinalIgnoreCase))
                return;

            // Load the document to be inserted.
            Document srcDoc = new Document(_insertDocPath);

            // Insert the document at the merge field location, preserving its formatting.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.DocumentFieldName);
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Suppress the default merge field text.
            args.Text = string.Empty;
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
