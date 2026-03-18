using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD for the product name.
        builder.InsertField("MERGEFIELD ProductName");
        builder.Writeln(); // Move to the next line.

        // Insert an INCLUDEPICTURE field whose source is taken from the MERGEFIELD ImagePath.
        // The \d switch tells Word to preserve the original image dimensions.
        // NOTE: Use a verbatim string and double the double‑quotes inside the field code.
        builder.InsertField(@"INCLUDEPICTURE ""{MERGEFIELD ImagePath}"" \\d");
        builder.Writeln(); // Separate entries.

        // Build a data source with product names and corresponding image file paths.
        DataTable products = new DataTable("Products");
        products.Columns.Add("ProductName", typeof(string));
        products.Columns.Add("ImagePath", typeof(string));

        // Example rows – replace with your actual data source.
        products.Rows.Add("Apple", @"C:\Images\apple.jpg");
        products.Rows.Add("Banana", @"C:\Images\banana.png");
        products.Rows.Add("Cherry", @"C:\Images\cherry.png");

        // Perform the mail merge. This fills the MERGEFIELDs with data from the table.
        doc.MailMerge.Execute(products);

        // After merging, update all fields so that the INCLUDEPICTURE fields load the images.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Products.docx");
    }
}
