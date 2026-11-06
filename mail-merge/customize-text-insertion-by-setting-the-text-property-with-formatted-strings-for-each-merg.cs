using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert merge fields into the template.
        builder.Writeln("Customer Name: ");
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln();

        builder.Writeln("Order Date: ");
        builder.InsertField("MERGEFIELD Date");
        builder.Writeln();

        builder.Writeln("Order Total: ");
        builder.InsertField("MERGEFIELD Amount");
        builder.Writeln();

        // Set up a custom field merging callback to format inserted text.
        doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

        // Prepare data for the merge.
        string[] fieldNames = { "Name", "Date", "Amount" };
        DateTime orderDate = new DateTime(2023, 12, 31);
        double orderTotal = 1234.56;
        object[] fieldValues = { "John Doe", orderDate, orderTotal };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        const string outputPath = "MergedDocument.docx";
        doc.Save(outputPath);

        // Validate that the merged content contains the expected formatted values.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("John Doe") ||
            !mergedText.Contains(orderDate.ToString("yyyy-MM-dd")) ||
            !mergedText.Contains(string.Format("{0:C}", orderTotal)))
        {
            throw new Exception("Mail merge validation failed: expected values not found in the document.");
        }
    }

    // Custom callback to format field values during mail merge.
    private class CustomFieldMergingCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            if (args.FieldValue == null)
            {
                args.Text = string.Empty;
                return;
            }

            switch (args.FieldName)
            {
                case "Date":
                    // Format DateTime fields as yyyy-MM-dd.
                    if (args.FieldValue is DateTime dt)
                        args.Text = dt.ToString("yyyy-MM-dd");
                    else
                        args.Text = args.FieldValue.ToString();
                    break;

                case "Amount":
                    // Format numeric fields as currency.
                    if (args.FieldValue is IFormattable fmt)
                        args.Text = fmt.ToString("C", System.Globalization.CultureInfo.CurrentCulture);
                    else
                        args.Text = args.FieldValue.ToString();
                    break;

                default:
                    // Default handling for other fields.
                    args.Text = args.FieldValue.ToString();
                    break;
            }
        }

        // Image fields are not used in this example.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
