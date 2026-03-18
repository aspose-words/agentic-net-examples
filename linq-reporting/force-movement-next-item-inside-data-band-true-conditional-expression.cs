using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare a data source with three rows.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Title");
            table.Columns.Add("FirstName");
            table.Columns.Add("LastName");
            table.Rows.Add("Mr.", "John", "Doe");
            table.Rows.Add("Mrs.", "Jane", "Smith");
            table.Rows.Add("Dr.", "Emily", "Brown");

            // Insert merge fields for the first record.
            InsertMergeFields(builder, "Row 1: ");

            // Insert a NEXTIF field whose condition is always true.
            // When the condition evaluates to true, the NEXTIF field forces the mail merge
            // engine to move to the next data record before processing subsequent merge fields.
            FieldNextIf nextIf = (FieldNextIf)builder.InsertField(FieldType.FieldNextIf, true);
            nextIf.LeftExpression = "1";
            nextIf.RightExpression = "1";
            nextIf.ComparisonOperator = "="; // 1 = 1 → true

            // Insert merge fields for the second record (will be merged after the NEXTIF moves to the next row).
            InsertMergeFields(builder, "Row 2: ");

            // Execute the mail merge.
            doc.MailMerge.Execute(table);

            // Save the resulting document.
            doc.Save("NextIf_TrueCondition.docx");
        }

        // Helper method to insert three merge fields (Title, FirstName, LastName) with optional preceding text.
        private static void InsertMergeFields(DocumentBuilder builder, string prefix)
        {
            // Optional text before the first merge field.
            if (!string.IsNullOrEmpty(prefix))
                builder.Write(prefix);

            // Insert the three merge fields.
            InsertMergeField(builder, "Title", " ", " ");
            InsertMergeField(builder, "FirstName", null, " ");
            InsertMergeField(builder, "LastName", null, null);

            // End the line.
            builder.Writeln();
        }

        // Helper method to insert a single MERGEFIELD with optional text before and after it.
        private static void InsertMergeField(DocumentBuilder builder, string fieldName, string textBefore, string textAfter)
        {
            if (!string.IsNullOrEmpty(textBefore))
                builder.Write(textBefore);

            FieldMergeField mergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            mergeField.FieldName = fieldName;

            if (!string.IsNullOrEmpty(textAfter))
                builder.Write(textAfter);
        }
    }
}
