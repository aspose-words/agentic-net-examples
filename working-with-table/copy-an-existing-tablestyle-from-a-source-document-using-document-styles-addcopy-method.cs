using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document and add a custom table style.
        Document sourceDoc = new Document();
        // Add a new table style with the specified name.
        Style customStyle = sourceDoc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // (Optional) Set any style properties that are available in the current version.
        // For this example we keep the style minimal; the focus is on copying it.

        // Save the source document (required for the example to work independently).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Create a destination document.
        Document destDoc = new Document();

        // Copy the custom table style from the source document to the destination document.
        Style copiedStyle = destDoc.Styles.AddCopy(sourceDoc.Styles["MyCustomTableStyle"]);

        // Verify that the style was copied successfully.
        if (copiedStyle == null || destDoc.Styles["MyCustomTableStyle"] == null)
        {
            throw new InvalidOperationException("The table style was not copied to the destination document.");
        }

        // Save the destination document.
        string destPath = "Destination.docx";
        destDoc.Save(destPath);

        // Simple validation that the output files exist.
        if (!File.Exists(sourcePath) || !File.Exists(destPath))
        {
            throw new FileNotFoundException("One of the output files was not created.");
        }

        // Indicate successful completion.
        Console.WriteLine("Table style copied successfully.");
    }
}
