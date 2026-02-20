using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Vba;

class VbaDocumentProcessing
{
    static void Main()
    {
        // Path to the original DOCX file that contains VBA macros.
        string originalPath = @"C:\Docs\Original.docx";

        // Load the original document.
        Document originalDoc = new Document(originalPath);

        // Clone the entire document, including its VBA project.
        Document clonedDoc = (Document)originalDoc.Clone(true);
        if (originalDoc.HasMacros)
        {
            // Clone the VBA project and assign it to the cloned document.
            clonedDoc.VbaProject = (VbaProject)originalDoc.VbaProject.Clone();
        }

        // Load additional documents that will be inserted/appended.
        Document docToInsert = new Document(@"C:\Docs\Insert.docx");
        Document docToAppend = new Document(@"C:\Docs\Append.docx");

        // Insert the document at the beginning of the cloned document.
        // Use DocumentBuilder to position the cursor at the start.
        DocumentBuilder builder = new DocumentBuilder(clonedDoc);
        builder.MoveToDocumentStart();
        builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Append another document to the end of the cloned document.
        clonedDoc.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);

        // Save the modified cloned document.
        string clonedPath = @"C:\Docs\ClonedModified.docx";
        clonedDoc.Save(clonedPath);

        // ------------------------------------------------------------
        // Split the original document into separate HTML files, one per section.
        // ------------------------------------------------------------
        HtmlSaveOptions splitOptions = new HtmlSaveOptions
        {
            // Split at each section break.
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            // Optional: keep the original file name as a base for split parts.
            ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection
        };

        // Save the split parts. The base file name is used; Aspose.Words will
        // generate files like Original.docx.part1.html, Original.docx.part2.html, etc.
        string splitBasePath = @"C:\Docs\Original_Split.html";
        originalDoc.Save(splitBasePath, splitOptions);
    }
}
