using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    // Path to the static logo image that will be inserted into every merged document.
    private const string LogoPath = "Logo.png";

    public static void Main()
    {
        if (!File.Exists(LogoPath))
        {
            // Create a simple placeholder file (not a real image, but sufficient for the demo)
            File.WriteAllBytes(LogoPath, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z8BQDwAF/wJ/lKXUAAAAAElFTkSuQmCC")); // PNG header bytes
        }
        // Create a simple mail‑merge template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a text merge field.
        builder.InsertField("MERGEFIELD Name", "<Name>");
        builder.Writeln();

        // Insert an image merge field. The field name after the "Image:" prefix is arbitrary;
        // the callback will supply the actual image.
        builder.InsertField("MERGEFIELD Image:Logo", "<Logo>");

        // Prepare a data source. The "Logo" column value is ignored because the callback
        // always uses the static image file.
        DataTable data = new DataTable("Employees");
        data.Columns.Add("Name", typeof(string));
        data.Columns.Add("Logo", typeof(string)); // placeholder column
        data.Rows.Add("Alice Johnson", "placeholder");
        data.Rows.Add("Bob Smith", "placeholder");

        // Assign the callback that supplies the static logo image.
        template.MailMerge.FieldMergingCallback = new StaticLogoCallback(LogoPath);

        // Perform the mail merge.
        template.MailMerge.Execute(data);

        // Save the merged document.
        template.Save("MergedDocument.docx");
    }

    // Callback that provides a static image for every image merge field.
    private class StaticLogoCallback : IFieldMergingCallback
    {
        private readonly string _imagePath;

        public StaticLogoCallback(string imagePath)
        {
            _imagePath = imagePath;
        }

        // No custom handling for text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Intentionally left blank.
        }

        // Called when an image merge field is encountered.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Use the ImageFileName property which works across .NET versions.
            args.ImageFileName = _imagePath;
        }
    }
}
