using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

public class Program
{
    public static void Main()
    {
        // Create a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create sample image files.
        // -----------------------------------------------------------------
        // A tiny red dot PNG (1x1 pixel) encoded in Base64.
        string redDotBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9W6XK2cAAAAASUVORK5CYII=";
        byte[] redDotBytes = Convert.FromBase64String(redDotBase64);
        string imagePath1 = Path.Combine(workDir, "red.png");
        File.WriteAllBytes(imagePath1, redDotBytes);

        // A tiny green dot PNG (1x1 pixel) encoded in Base64.
        string greenDotBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwADhgGk2cVYVwAAAABJRU5ErkJggg==";
        byte[] greenDotBytes = Convert.FromBase64String(greenDotBase64);
        string imagePath2 = Path.Combine(workDir, "green.png");
        File.WriteAllBytes(imagePath2, greenDotBytes);

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table that will be repeated for each product.
        Table table = builder.StartTable();

        // First cell – product name.
        builder.InsertCell();
        builder.Write("<<[p.Name]>>");

        // Second cell – image placeholder inside a textbox.
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag that references the ImagePath property.
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(workDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Red Dot", ImagePath = imagePath1 },
                new Product { Name = "Green Dot", ImagePath = imagePath2 }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        report.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = "";
    public string ImagePath { get; set; } = "";
}
