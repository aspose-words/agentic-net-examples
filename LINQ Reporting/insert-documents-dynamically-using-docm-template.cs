using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCM template. The template must contain a bookmark named "InsertHere"
        // where the dynamic documents will be placed.
        Document template = new Document("Template.docm");

        // OPTIONAL: Populate merge fields in the template using ReportingEngine.
        // This step demonstrates how you can fill the template before inserting documents.
        var reportData = new
        {
            Title = "Dynamic Report",
            GeneratedOn = DateTime.Today.ToString("d")
        };
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, reportData, "ds");

        // List of source documents that will be inserted into the template.
        // These can be any supported format (DOCX, DOC, RTF, etc.).
        List<string> sourceFiles = new List<string>
        {
            "Section1.docx",
            "Section2.docx",
            "Section3.docx"
        };

        // Create a DocumentBuilder attached to the template.
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert each source document at the bookmark location.
        foreach (string filePath in sourceFiles)
        {
            // Load the source document.
            Document sourceDoc = new Document(filePath);

            // Move the cursor to the bookmark where the insertion should occur.
            builder.MoveToBookmark("InsertHere");

            // Insert the source document while keeping its original formatting.
            // InsertDocument mimics Word's copy‑paste behavior.
            builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

            // After insertion the bookmark is consumed. Re‑create it at the end of the
            // newly inserted content so that the next document can be placed after it.
            // This is done by moving to the end of the inserted content and adding a new bookmark.
            builder.MoveToDocumentEnd();
            builder.StartBookmark("InsertHere");
            builder.EndBookmark("InsertHere");
        }

        // Save the final document. Saving as DOCX removes macros; save as DOCM if you need to keep them.
        template.Save("Result.docx");
    }
}
