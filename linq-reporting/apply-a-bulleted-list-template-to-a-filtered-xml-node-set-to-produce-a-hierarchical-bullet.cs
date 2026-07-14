using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample XML data.
        // -----------------------------------------------------------------
        const string xmlFile = "data.xml";
        File.WriteAllText(xmlFile,
@"<Catalog>
    <Category name='Fruits' active='true'>
        <SubCategory name='Apple' />
        <SubCategory name='Banana' />
    </Category>
    <Category name='Vegetables' active='false'>
        <SubCategory name='Carrot' />
        <SubCategory name='Lettuce' />
    </Category>
    <Category name='Beverages' active='true'>
        <SubCategory name='Coffee' />
        <SubCategory name='Tea' />
    </Category>
</Catalog>");

        // -----------------------------------------------------------------
        // 2. Load XML and build a filtered model (only active categories).
        // -----------------------------------------------------------------
        ReportModel model = new();
        XDocument doc = XDocument.Load(xmlFile);
        foreach (XElement catElem in doc.Root!.Elements("Category"))
        {
            if (bool.TryParse(catElem.Attribute("active")?.Value, out bool isActive) && isActive)
            {
                var category = new Category
                {
                    Name = catElem.Attribute("name")?.Value ?? string.Empty
                };

                foreach (XElement subElem in catElem.Elements("SubCategory"))
                {
                    category.SubCategories.Add(new SubCategory
                    {
                        Name = subElem.Attribute("name")?.Value ?? string.Empty
                    });
                }

                model.Categories.Add(category);
            }
        }

        // -----------------------------------------------------------------
        // 3. Create the LINQ Reporting template.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Create a bullet list that will be used for both levels.
        List bulletList = template.Lists.Add(ListTemplate.BulletDefault);

        // Begin outer foreach – categories.
        builder.Writeln("<<foreach [cat in Categories]>>");

        // Category level (list level 0).
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[cat.Name]>>");

        // Begin inner foreach – sub‑categories.
        builder.Writeln("<<foreach [sub in cat.SubCategories]>>");

        // Sub‑category level (list level 1).
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[sub.Name]>>");

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the result.
        // -----------------------------------------------------------------
        const string outputFile = "Report.docx";
        template.Save(outputFile);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputFile)}");
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, with property initializers to avoid warnings)
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<SubCategory> SubCategories { get; set; } = new();
}

public class SubCategory
{
    public string Name { get; set; } = string.Empty;
}
