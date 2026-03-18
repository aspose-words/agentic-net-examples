using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

namespace RestartListNumberingExample
{
    // Implements a callback that inserts a LISTNUM field before each group.
    // The LISTNUM field will start counting from 1, effectively restarting the list.
    public class RestartNumberCallback : IFieldMergingCallback
    {
        // Called for every merge field encountered during MailMerge.Execute.
        public void FieldMerging(FieldMergingArgs args)
        {
            // The template should contain a placeholder <<restartNum>> placed before the <<foreach>> region.
            if (args.FieldName.Equals("restartNum", StringComparison.OrdinalIgnoreCase))
            {
                // Move the document builder to the location of the <<restartNum>> field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.FieldName);

                // Insert a LISTNUM field that restarts numbering at 1.
                // The field will be placed where the <<restartNum>> tag was.
                FieldListNum listNum = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);
                listNum.StartingNumber = "1";

                // Remove the original placeholder text (the merge field itself) so it does not appear in the output.
                args.Text = string.Empty;
            }
        }

        // Not used in this example, but required by the interface.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    class Program
    {
        static void Main()
        {
            // Load the Word template that contains the <<restartNum>> tag before the <<foreach>> region.
            Document doc = new Document("Template.docx");

            // Ensure that list numbering does NOT automatically restart at each section,
            // because we will control the restart manually via the LISTNUM field.
            doc.MailMerge.RestartListsAtEachSection = false;

            // Assign the custom callback that will replace <<restartNum>> with a LISTNUM field.
            doc.MailMerge.FieldMergingCallback = new RestartNumberCallback();

            // Prepare sample data for the mail merge.
            // The data source must contain a field named "Group" (or any name you use in the template)
            // and the template should have a <<foreach>> region that repeats for each group.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Name");
            table.Columns.Add("Department");

            // Add rows – each change in Department will be a new group.
            table.Rows.Add("John Doe", "Sales");
            table.Rows.Add("Jane Smith", "Sales");
            table.Rows.Add("Bob Johnson", "Marketing");
            table.Rows.Add("Alice Brown", "Marketing");
            table.Rows.Add("Tom White", "HR");

            // Execute mail merge with the data source.
            // The <<restartNum>> field will be processed for each group, restarting the list.
            doc.MailMerge.Execute(table);

            // Save the resulting document.
            doc.Save("Result.docx");
        }
    }
}
