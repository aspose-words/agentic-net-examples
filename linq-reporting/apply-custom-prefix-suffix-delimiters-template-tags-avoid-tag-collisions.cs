using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load a template that contains default mustache tags like {{FirstName}}.
        Document doc = new Document("Template.docx");

        // Allow mail merge to work with plain‑text tags.
        doc.MailMerge.UseNonMergeFields = true;

        // Define custom delimiters to avoid collisions with other tags.
        const string customStart = "[[";
        const string customEnd = "]]";

        // Replace the default start delimiter "{{" with the custom one.
        doc.Range.Replace("{{", customStart);
        // Replace the default end delimiter "}}" with the custom one.
        doc.Range.Replace("}}", customEnd);

        // Create a simple data source for the mail merge.
        DataTable data = new DataTable("Employees");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Rows.Add("John", "Doe");
        data.Rows.Add("Jane", "Smith");

        // Execute the mail merge. The tags now use the custom delimiters.
        doc.MailMerge.Execute(data);

        // Save the merged document.
        doc.Save("Result.docx");
    }
}
