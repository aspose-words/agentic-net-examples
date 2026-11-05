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

        // Insert a static Table of Contents field at the beginning of the document.
        // The TOC will be updated after the mail merge execution.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a line break after the TOC.

        // Insert a heading that will contain a merge field.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("Report for ");
        builder.InsertField("MERGEFIELD CustomerName \\* MERGEFORMAT");
        builder.Writeln();

        // Insert a normal paragraph with another merge field.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Write("Date: ");
        builder.InsertField("MERGEFIELD ReportDate \\* MERGEFORMAT");
        builder.Writeln();

        // Prepare simple mail merge data.
        string[] fieldNames = new string[] { "CustomerName", "ReportDate" };
        object[] fieldValues = new object[] { "Acme Corp", DateTime.Now.ToString("yyyy-MM-dd") };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Validate that merge fields were replaced.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("Acme Corp"))
        {
            throw new InvalidOperationException("Mail merge failed: CustomerName not replaced.");
        }

        // Update all fields (including the TOC) to reflect the merged content.
        doc.UpdateFields();

        // Save the merged document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
