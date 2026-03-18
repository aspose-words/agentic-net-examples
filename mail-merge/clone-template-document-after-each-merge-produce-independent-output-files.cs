using System;
using System.Data;
using Aspose.Words;

class MailMergeCloneExample
{
    static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Load the template document that contains MERGEFIELDs.
        Document template = new Document("Template.docx");

        // Create a data source with several rows.
        DataTable data = new DataTable("Employees");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Title");
        data.Rows.Add("John", "Doe", "Manager");
        data.Rows.Add("Jane", "Smith", "Developer");
        data.Rows.Add("Bob", "Brown", "Analyst");

        // For each row create an independent document, merge, and save.
        int docIndex = 1;
        foreach (DataRow row in data.Rows)
        {
            // Deep clone the template so changes do not affect other outputs.
            Document mergedDoc = (Document)template.Clone();

            // Perform mail merge for the current record.
            mergedDoc.MailMerge.Execute(row);

            // Save the result to a unique file name.
            string outputPath = $"MergedDocument_{docIndex}.docx";
            mergedDoc.Save(outputPath);
            docIndex++;
        }
    }
}
