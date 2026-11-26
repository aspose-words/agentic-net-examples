using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeCustomText
{
    // Custom callback to control how each merge field is inserted.
    public class CustomFieldMergingCallback : IFieldMergingCallback
    {
        // This method is called for every merge field during MailMerge.
        public void FieldMerging(FieldMergingArgs e)
        {
            // If the field value is not null, format it as a custom string.
            if (e.FieldValue != null)
            {
                // Example of custom formatting: wrap the value in brackets.
                e.Text = $"[Custom:{e.FieldValue}]";
            }
            else
            {
                // Provide a placeholder for missing values.
                e.Text = "[Missing]";
            }
        }

        // Not used in this example, but required by the interface.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple template with two merge fields.
            builder.Writeln("Dear ");
            builder.InsertField("MERGEFIELD FirstName");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName");
            builder.Writeln(",");

            builder.Writeln("Your account has been created successfully.");
            builder.Writeln("Thank you!");

            // Assign the custom field merging callback.
            doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

            // Prepare data for the merge.
            string[] fieldNames = { "FirstName", "LastName" };
            object[] fieldValues = { "John", "Doe" };

            // Execute the mail merge.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Validate that the merged document contains the expected custom text.
            string mergedText = doc.GetText();
            if (!mergedText.Contains("[Custom:John]") || !mergedText.Contains("[Custom:Doe]"))
            {
                throw new InvalidOperationException("Merged document does not contain the expected custom field values.");
            }

            // Optionally clean up any unused fields (none in this example).
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

            // Save the merged document.
            const string outputPath = "MergedDocument.docx";
            doc.Save(outputPath);
        }
    }
}
