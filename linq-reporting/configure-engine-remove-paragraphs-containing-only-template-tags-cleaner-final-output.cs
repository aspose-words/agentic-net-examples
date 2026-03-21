using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

class RemoveTemplateTagParagraphs
{
    static void Main()
    {
        // Create a new document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs with MERGEFIELD tags.
        builder.Writeln("Hello <<FirstName>> <<LastName>>.");
        builder.Writeln("Address: <<Address>>");
        // This paragraph contains only a MERGEFIELD tag and will be removed after cleanup.
        builder.Writeln("<<OnlyTag>>");

        // Configure mail merge to remove paragraphs that become empty after the merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // Example data for the mail merge.
        string[] fieldNames = { "FirstName", "LastName", "Address", "OnlyTag" };
        object[] fieldValues = { "John", "Doe", "123 Main St.", "" };

        // Perform the mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the cleaned‑up document to the current directory.
        string outputPath = "ResultWithoutEmptyTagParagraphs.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to {outputPath}");
    }
}
