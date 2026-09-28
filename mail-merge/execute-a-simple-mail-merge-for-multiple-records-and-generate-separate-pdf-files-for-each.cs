using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear <<Name>>,");
        builder.Writeln("Your address is <<Address>>.");
        builder.Writeln("Thank you for your business.");
        builder.Writeln("Sincerely,");
        builder.Writeln("Company XYZ");

        // Prepare data for mail merge.
        DataTable data = new DataTable("Customers");
        data.Columns.Add("Name", typeof(string));
        data.Columns.Add("Address", typeof(string));

        data.Rows.Add("Alice Johnson", "123 Maple Street");
        data.Rows.Add("Bob Smith", "456 Oak Avenue");
        data.Rows.Add("Carol Davis", "789 Pine Road");

        // Field names used in the template.
        string[] fieldNames = { "Name", "Address" };

        // Generate a separate PDF for each record.
        int index = 1;
        foreach (DataRow row in data.Rows)
        {
            // Clone the template to keep it unchanged for the next iteration.
            Document doc = (Document)template.Clone(true);

            // Execute mail merge for the current record.
            object[] fieldValues = { row["Name"], row["Address"] };
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the result as a PDF file.
            string fileName = $"Customer_{index}_{row["Name"].ToString().Replace(' ', '_')}.pdf";
            doc.Save(fileName, SaveFormat.Pdf);

            index++;
        }
    }
}
