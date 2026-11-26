using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a merge field named "Name".
        builder.Font.Size = 12;
        builder.InsertField("MERGEFIELD Name");

        // Prepare simple mail‑merge data.
        string[] fieldNames = { "Name" };
        object[] fieldValues = { "John Doe" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // After merging, locate the run that contains the merged name and apply bold formatting.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "John Doe")
            {
                run.Font.Bold = true;
            }
        }

        // Save the merged document.
        string outputPath = "Merged.docx";
        doc.Save(outputPath);

        // Validation: ensure the merged name exists and is bold.
        bool boldApplied = false;
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "John Doe" && run.Font.Bold)
            {
                boldApplied = true;
                break;
            }
        }

        if (!boldApplied)
        {
            throw new InvalidOperationException("Bold formatting was not applied to the merged name.");
        }

        // Confirm that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The merged document was not created.", outputPath);
        }

        Console.WriteLine($"Mail merge completed successfully. Output saved to \"{outputPath}\".");
    }
}
