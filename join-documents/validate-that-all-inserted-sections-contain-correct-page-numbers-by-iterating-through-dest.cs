using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class JoinDocumentsValidatePageNumbers
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -------------------- Create source document 1 --------------------
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);

        // Insert a header with a PAGE field.
        builder1.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder1.InsertField(FieldType.FieldPage, true);

        // Add content that spans two pages.
        builder1.MoveToDocumentStart();
        builder1.Writeln("Source Document 1 - Page 1");
        builder1.InsertBreak(BreakType.PageBreak);
        builder1.Writeln("Source Document 1 - Page 2");

        string srcPath1 = Path.Combine(artifactsDir, "Source1.docx");
        srcDoc1.Save(srcPath1, SaveFormat.Docx);

        // -------------------- Create source document 2 --------------------
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);

        // Insert a header with a PAGE field.
        builder2.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder2.InsertField(FieldType.FieldPage, true);

        // Add content that spans three pages.
        builder2.MoveToDocumentStart();
        builder2.Writeln("Source Document 2 - Page 1");
        builder2.InsertBreak(BreakType.PageBreak);
        builder2.Writeln("Source Document 2 - Page 2");
        builder2.InsertBreak(BreakType.PageBreak);
        builder2.Writeln("Source Document 2 - Page 3");

        string srcPath2 = Path.Combine(artifactsDir, "Source2.docx");
        srcDoc2.Save(srcPath2, SaveFormat.Docx);

        // -------------------- Create destination document --------------------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);

        // Insert a header with a PAGE field for the destination document.
        dstBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        dstBuilder.InsertField(FieldType.FieldPage, true);

        // Add some initial content.
        dstBuilder.MoveToDocumentStart();
        dstBuilder.Writeln("Destination Document - Start");
        dstBuilder.InsertBreak(BreakType.PageBreak);
        dstBuilder.Writeln("Destination Document - End");

        // -------------------- Append source documents --------------------
        // Append first source document.
        dstDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);
        // Append second source document.
        dstDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Ensure layout and fields are up‑to‑date.
        dstDoc.UpdatePageLayout();
        dstDoc.UpdateFields();

        // -------------------- Validation of page numbers --------------------
        // After appending, each section should contain a PAGE field whose result reflects the actual page number.
        // Iterate through all sections and verify the PAGE field result is a non‑empty numeric string.
        for (int i = 0; i < dstDoc.Sections.Count; i++)
        {
            Section section = dstDoc.Sections[i];
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];

            // If the section has no primary header, try any header.
            if (header == null)
                header = section.HeadersFooters[HeaderFooterType.HeaderFirst] ??
                         section.HeadersFooters[HeaderFooterType.HeaderEven];

            if (header == null)
                throw new InvalidOperationException($"Section {i + 1} does not contain a header.");

            // Find the first PAGE field in the header.
            Field pageField = null;
            foreach (Field field in header.Range.Fields)
            {
                if (field.Type == FieldType.FieldPage)
                {
                    pageField = field;
                    break;
                }
            }

            if (pageField == null)
                throw new InvalidOperationException($"Section {i + 1} header does not contain a PAGE field.");

            // The field result should be a number representing the page.
            string result = pageField.Result?.Trim();
            if (string.IsNullOrEmpty(result) || !int.TryParse(result, out _))
                throw new InvalidOperationException($"Section {i + 1} PAGE field result is invalid: '{result}'.");
        }

        // -------------------- Save merged document --------------------
        string mergedPath = Path.Combine(artifactsDir, "MergedDocument.docx");
        dstDoc.Save(mergedPath, SaveFormat.Docx);

        // Verify that the merged file exists.
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not saved correctly.", mergedPath);

        // Indicate successful completion.
        Console.WriteLine("Document merged and page numbers validated successfully.");
    }
}
