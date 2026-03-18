using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class RemoveTemplateTagParagraphs
{
    static void Main()
    {
        // Path to the source template document that contains MERGEFIELD tags.
        string inputPath = @"C:\Docs\TemplateWithTags.docx";

        // Load the document.
        Document doc = new Document(inputPath);

        // Configure mail merge to remove paragraphs that become empty after the merge.
        // This will delete paragraphs that contained only MERGEFIELD tags (or punctuation marks).
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        // Optional: treat paragraphs that contain only punctuation as empty as well.
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // Example data for the mail merge.
        // Field names must match the MERGEFIELD names in the template.
        string[] fieldNames = { "FirstName", "LastName", "Address" };
        object[] fieldValues = { "John", "Doe", "123 Main St." };

        // Perform the mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the cleaned‑up document.
        string outputPath = @"C:\Docs\ResultWithoutEmptyTagParagraphs.docx";
        doc.Save(outputPath);
    }
}
