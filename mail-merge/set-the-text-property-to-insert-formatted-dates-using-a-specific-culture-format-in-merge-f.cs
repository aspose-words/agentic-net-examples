using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a merge field named "ReportDate".
        builder.Writeln("Report Date: ");
        builder.InsertField("MERGEFIELD ReportDate \\* MERGEFORMAT");

        // Prepare data for the merge.
        string[] fieldNames = { "ReportDate" };
        DateTime currentDate = DateTime.Now;
        // Format the date using a specific culture (e.g., German - Germany).
        string formattedDate = currentDate.ToString("D", new CultureInfo("de-DE"));
        object[] fieldValues = { formattedDate };

        // Execute the mail merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");
        doc.Save(outputPath);

        // Validate that the merged document contains the formatted date.
        string docText = doc.GetText();
        if (!docText.Contains(formattedDate))
        {
            throw new InvalidOperationException("The merged document does not contain the expected formatted date.");
        }

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The merged document file was not created.", outputPath);
        }
    }
}
