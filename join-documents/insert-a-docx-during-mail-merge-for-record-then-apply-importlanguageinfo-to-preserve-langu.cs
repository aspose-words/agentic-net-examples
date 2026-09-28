using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a template DOCX that contains two merge fields:
        //    - Name : will be replaced with a person's name.
        //    - InsertDoc : placeholder where another document will be inserted.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        var templateDoc = new Document();
        var tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Record:");
        tmplBuilder.InsertField("MERGEFIELD Name \\* MERGEFORMAT");
        tmplBuilder.Writeln();
        tmplBuilder.InsertField("MERGEFIELD InsertDoc \\* MERGEFORMAT");
        tmplBuilder.Writeln();
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a DOCX that will be inserted during mail merge.
        //    Set its language (LocaleId) to English (United States) = 1033.
        // -----------------------------------------------------------------
        string insertPath = Path.Combine(outputDir, "Insert.docx");
        var insertDoc = new Document();
        var insBuilder = new DocumentBuilder(insertDoc);
        insBuilder.Font.LocaleId = 1033; // English (United States)
        insBuilder.Writeln("This is the inserted document with language settings preserved.");
        insertDoc.Save(insertPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the created documents.
        // -----------------------------------------------------------------
        var template = new Document(templatePath);
        var insert = new Document(insertPath);

        // -----------------------------------------------------------------
        // 4. Prepare a simple data source for mail merge (two records).
        // -----------------------------------------------------------------
        var table = new DataTable("Records");
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");

        // Master document that will hold all merged records.
        var masterDoc = new Document();

        foreach (DataRow row in table.Rows)
        {
            // Clone the template for the current record.
            var recordDoc = (Document)template.Clone(true);

            // Execute mail merge for the Name field.
            recordDoc.MailMerge.Execute(new[] { "Name" }, new object[] { row["Name"] });

            // Locate the InsertDoc merge field.
            var insertField = recordDoc.Range.Fields
                .FirstOrDefault(f => f.Type == FieldType.FieldMergeField &&
                                     (f as FieldMergeField)?.FieldName == "InsertDoc");

            if (insertField != null)
            {
                // Move the builder to the start of the merge field.
                var builder = new DocumentBuilder(recordDoc);
                builder.MoveTo(insertField.Start);

                // Insert the document at the field location, keeping source formatting.
                builder.InsertDocument(insert, ImportFormatMode.KeepSourceFormatting);

                // Remove the placeholder merge field.
                insertField.Remove();
            }

            // Append the processed record document to the master document.
            masterDoc.AppendDocument(recordDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // -----------------------------------------------------------------
        // 5. Save the final merged document as PDF.
        // -----------------------------------------------------------------
        string outputPdf = Path.Combine(outputDir, "MergedOutput.pdf");
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("The merged PDF was not created.");
        }

        Console.WriteLine($"Merged PDF created successfully at: {outputPdf}");
    }
}
