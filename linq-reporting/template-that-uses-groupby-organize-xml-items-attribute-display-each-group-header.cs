// Load XML, group items by an attribute, and generate a Word report using Aspose.Words Reporting Engine.
// The code follows the required lifecycle: create a Document, build a template, populate it, and save the result.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsGroupByExample
{
    // Simple POCO representing a single XML item.
    public class Item
    {
        public string Name { get; set; }      // Example element value.
        public string Category { get; set; }  // Attribute used for grouping.
    }

    // POCO representing a group of items that share the same attribute value.
    public class ItemGroup
    {
        public string Category { get; set; }          // Group header (the attribute value).
        public List<Item> Items { get; set; }         // Items belonging to this group.
    }

    class Program
    {
        static void Main()
        {
            // ---------- 1. Load and parse the XML ----------
            // Example XML structure:
            // <catalog>
            //   <product category="Books"><name>Everyday Italian</name></product>
            //   <product category="Books"><name>The C Programming Language</name></product>
            //   <product category="Electronics"><name>Smartphone X</name></product>
            // </catalog>

            string xmlPath = "catalog.xml"; // Path to your XML file.
            XDocument xDoc = XDocument.Load(xmlPath);

            // Extract items from XML.
            List<Item> items = xDoc.Root
                .Elements("product")
                .Select(p => new Item
                {
                    Name = (string)p.Element("name"),
                    Category = (string)p.Attribute("category")
                })
                .ToList();

            // ---------- 2. Group items by the 'category' attribute ----------
            List<ItemGroup> groups = items
                .GroupBy(i => i.Category)
                .Select(g => new ItemGroup
                {
                    Category = g.Key,
                    Items = g.ToList()
                })
                .ToList();

            // ---------- 3. Create a new Word document (template) ----------
            Document doc = new Document();                     // create
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the reporting template using Aspose.Words LINQ Reporting Engine syntax.
            // The outer loop iterates over each group, displaying the group header.
            // The inner loop iterates over the items belonging to the current group.
            builder.Writeln("<<foreach [group]>>");            // start group loop
            builder.Writeln("Category: <<[group.Category]>>"); // group header
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item] in [group.Items]>>"); // start items loop
            builder.Writeln("- <<[item.Name]>>");
            builder.Writeln("<<endfor>>");                     // end items loop
            builder.Writeln("<<endfor>>");                     // end group loop

            // ---------- 4. Populate the template with the grouped data ----------
            ReportingEngine engine = new ReportingEngine();
            // The data source name "groups" is used in the template (optional here because we reference the object directly).
            engine.BuildReport(doc, groups, "groups");

            // ---------- 5. Save the final document ----------
            doc.Save("GroupedReport.docx");                    // save
        }
    }
}
