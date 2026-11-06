using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
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

        // Prepare data for the mail merge.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Validate that the merge replaced the fields.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("John") || !mergedText.Contains("Doe"))
        {
            throw new InvalidOperationException("Mail merge did not replace all fields correctly.");
        }

        // Save the merged document as PDF.
        string outputPath = "MergedOutput.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Ensure the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The PDF file was not created.", outputPath);
        }
    }
}
