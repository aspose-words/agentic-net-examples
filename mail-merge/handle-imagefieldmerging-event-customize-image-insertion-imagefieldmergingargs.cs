using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

class CustomImageHandler : IFieldMergingCallback
{
    // No custom text merging needed.
    void IFieldMergingCallback.FieldMerging(FieldMergingArgs args) { }

    // Called for MERGEFIELDS whose name starts with "Image:".
    void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
    {
        // The data source provides a short identifier for the image.
        string key = args.FieldValue?.ToString();
        if (string.IsNullOrEmpty(key))
            return;

        // Resolve the identifier to a physical file path.
        if (ImageMap.TryGetValue(key, out string filePath) && File.Exists(filePath))
        {
#if NET461_OR_GREATER || JAVA
            // For .NET Framework or Java, use System.Drawing.Image.
            args.Image = Image.FromFile(filePath);
#elif NET5_0_OR_GREATER
            // For .NET 5+ use SkiaSharp (Aspose.Words supports SKBitmap as an image source).
            args.Image = SkiaSharp.SKBitmap.Decode(filePath);
            // Also set the filename for completeness.
            args.ImageFileName = filePath;
#endif
        }
    }

    // Simple lookup table mapping short names to actual image files.
    private static readonly Dictionary<string, string> ImageMap = new()
    {
        { "Logo", @"C:\Images\Logo.jpg" },
        { "Badge", @"C:\Images\Badge.png" }
    };
}

class Program
{
    static void Main()
    {
        // Load a template that contains a field like MERGEFIELD Image:Photo.
        Document doc = new Document(@"C:\Templates\Report.docx");

        // Attach the custom callback that supplies images.
        doc.MailMerge.FieldMergingCallback = new CustomImageHandler();

        // Build a data source. The "Photo" column holds the short identifiers used above.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name");
        table.Columns.Add("Photo");
        table.Rows.Add("John Doe", "Logo");
        table.Rows.Add("Jane Smith", "Badge");

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the merged document.
        doc.Save(@"C:\Output\ReportMerged.docx");
    }
}
