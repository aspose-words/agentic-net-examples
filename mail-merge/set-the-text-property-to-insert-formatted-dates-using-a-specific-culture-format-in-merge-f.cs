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

        // Insert a merge field named "Date".
        builder.Writeln("Report generated on:");
        builder.InsertField("MERGEFIELD Date");
        builder.Writeln(); // Add a line break.

        // Prepare the data source.
        string[] fieldNames = { "Date" };
        // Use a specific culture (e.g., French - France) to format the date.
        CultureInfo culture = new CultureInfo("fr-FR");
        DateTime now = DateTime.Now;
        string formattedDate = now.ToString("D", culture); // Long date pattern for the culture.
        object[] fieldValues = { formattedDate };

        // Execute the mail merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Validate that the merge succeeded.
        string mergedText = doc.GetText();
        if (!mergedText.Contains(formattedDate))
        {
            throw new InvalidOperationException("The formatted date was not inserted into the document.");
        }

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.docx");
        doc.Save(outputPath);
    }
}
