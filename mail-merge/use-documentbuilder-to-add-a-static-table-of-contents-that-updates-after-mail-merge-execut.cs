using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a static Table of Contents field.
        // The TOC will include headings level 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a blank line after TOC.

        // Insert a heading that will be populated by a merge field.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.InsertField("MERGEFIELD Title");
        builder.Writeln(); // End heading paragraph.

        // Insert a normal paragraph with another merge field.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.InsertField("MERGEFIELD Content");
        builder.Writeln();

        // Prepare simple mail‑merge data.
        string[] fieldNames = new string[] { "Title", "Content" };
        object[] fieldValues = new object[] { "Sample Document Title", "This is the body text inserted via mail merge." };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Update all fields (including the TOC) after the merge.
        doc.UpdateFields();

        // Define output path.
        string outputPath = "MergedDocument.docx";

        // Save the merged document.
        doc.Save(outputPath);

        // Validation: ensure the file was created and is not empty.
        FileInfo info = new FileInfo(outputPath);
        if (!info.Exists || info.Length == 0)
        {
            throw new InvalidOperationException("The merged document was not created correctly.");
        }

        // Indicate success.
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
