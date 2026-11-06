using System;
using System.IO;
using Aspose.Words;

public class MailMergeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple template with merge fields.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(",");
        builder.Writeln("Your order number is ");
        builder.InsertField("MERGEFIELD OrderNumber");
        builder.Writeln(".");

        // Define field names and corresponding values.
        string[] fieldNames = new[] { "FirstName", "LastName", "OrderNumber" };
        object[] fieldValues = new object[] { "John", "Doe", "12345" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Validate that the merge replaced all fields.
        string mergedText = doc.GetText();
        foreach (object value in fieldValues)
        {
            if (value is string str && !mergedText.Contains(str))
            {
                throw new InvalidOperationException($"Merged document does not contain expected value: {str}");
            }
        }

        // Save the merged document.
        string outputPath = "MergedResult.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The merged document was not saved correctly.", outputPath);
        }

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine($"Mail merge completed successfully. File saved to: {Path.GetFullPath(outputPath)}");
    }
}
