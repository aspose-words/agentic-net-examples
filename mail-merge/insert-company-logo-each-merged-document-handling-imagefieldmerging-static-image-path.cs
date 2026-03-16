using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace CompanyMailMerge
{
    // Callback that supplies a static logo image for any image merge field.
    public class LogoImageHandler : IFieldMergingCallback
    {
        private readonly string _logoPath;

        public LogoImageHandler(string logoPath)
        {
            _logoPath = logoPath;
        }

        // No custom handling for text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Intentionally left blank.
        }

        // Called when the mail merge engine encounters an image merge field.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Supply the static logo file name. The engine will load the image from this path.
            args.ImageFileName = _logoPath;

            // Optional: set explicit dimensions (points). Comment out if original size is desired.
            // args.ImageWidth = new MergeFieldImageDimension(100, MergeFieldImageDimensionUnit.Point);
            // args.ImageHeight = new MergeFieldImageDimension(50, MergeFieldImageDimensionUnit.Point);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the template document that contains MERGEFIELD Image:Logo (or any image field).
            const string templatePath = @"C:\Templates\LetterTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Assign the image merging callback with the static logo path.
            const string logoPath = @"C:\Images\CompanyLogo.png";
            doc.MailMerge.FieldMergingCallback = new LogoImageHandler(logoPath);

            // Prepare a simple data source. The actual fields are not important for the logo,
            // but at least one non‑image field is needed to trigger the merge.
            DataTable data = new DataTable("Employees");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Rows.Add("John", "Doe");
            data.Rows.Add("Jane", "Smith");

            // Execute the mail merge. The image field will be replaced by the static logo.
            doc.MailMerge.Execute(data);

            // Save the merged document.
            const string outputPath = @"C:\Output\MergedLetters.docx";
            doc.Save(outputPath);
        }
    }
}
