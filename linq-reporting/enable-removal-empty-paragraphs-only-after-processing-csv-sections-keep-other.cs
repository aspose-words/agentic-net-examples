using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

class RemoveEmptyParagraphsAfterCsvProcessing
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("InputDocument.docx");

        // Identify sections that contain CSV data.
        // For this example we assume such sections are marked with a bookmark whose name starts with "CSVSection".
        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            if (!bookmark.Name.StartsWith("CSVSection", StringComparison.OrdinalIgnoreCase))
                continue;

            // Get the section that contains the bookmark.
            Section originalSection = (Section)bookmark.BookmarkStart.GetAncestor(NodeType.Section);
            if (originalSection == null)
                continue;

            // Clone the section into a temporary document for isolated processing.
            Document tempDoc = new Document();
            Section clonedSection = (Section)tempDoc.ImportNode(originalSection, true);
            tempDoc.AppendChild(clonedSection);

            // Prepare CSV data as a DataTable.
            // In a real scenario you would parse the CSV file; here we create a simple table for demonstration.
            DataTable csvData = new DataTable("CsvTable");
            csvData.Columns.Add("Column1");
            csvData.Columns.Add("Column2");
            csvData.Rows.Add(new object[] { "Value A1", "Value A2" });
            csvData.Rows.Add(new object[] { "", "" });               // This row will generate an empty paragraph.
            csvData.Rows.Add(new object[] { "Value B1", "Value B2" });

            // Enable removal of empty paragraphs for this mail‑merge operation.
            tempDoc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            // Optional: treat paragraphs with only punctuation as empty.
            tempDoc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

            // Execute the mail merge on the temporary document.
            tempDoc.MailMerge.Execute(csvData);

            // Import the cleaned section back into the original document,
            // replacing the original section while leaving other sections untouched.
            NodeImporter importer = new NodeImporter(tempDoc, doc, ImportFormatMode.KeepSourceFormatting);
            Section cleanedSection = (Section)importer.ImportNode(clonedSection, true);

            int sectionIndex = doc.Sections.IndexOf(originalSection);
            doc.Sections.RemoveAt(sectionIndex);
            doc.Sections.Insert(sectionIndex, cleanedSection);
        }

        // Save the modified document.
        doc.Save("OutputDocument.docx", SaveFormat.Docx);
    }
}
