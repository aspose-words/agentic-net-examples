using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    // Callback to handle missing merge fields.
    private class MissingFieldHandler : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            // If the data source does not contain a value for the field, treat it as missing.
            if (args.FieldValue == null)
            {
                Console.WriteLine($"Missing field encountered: {args.FieldName}");
                // Provide a default replacement text.
                args.Text = "[Missing]";
            }
        }

        // Required by the interface but not used for image fields.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    public static void Main()
    {
        // Create a new document and build the template with merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD Name", "«Name»");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Age", "«Age»");
        builder.Writeln();

        // Attach the missing‑field handler.
        doc.MailMerge.FieldMergingCallback = new MissingFieldHandler();

        // Prepare data source – Age field is intentionally omitted.
        string[] fieldNames = { "Name" };
        object[] fieldValues = { "John Doe" };

        // Perform the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.docx");
        doc.Save(outputPath);

        // Load the result to verify the merge.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("John Doe"))
            throw new InvalidOperationException("The Name field was not merged correctly.");

        if (!resultText.Contains("[Missing]"))
            throw new InvalidOperationException("The missing Age field was not handled correctly.");

        Console.WriteLine("Mail merge completed successfully. Output saved to: " + outputPath);
    }
}
