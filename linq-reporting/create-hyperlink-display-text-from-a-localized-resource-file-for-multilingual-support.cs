using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure code pages are available (required by Aspose.Words in some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create simple resource files for default (en) and French (fr) cultures.
        CreateResourceFiles();

        // Set the current culture to demonstrate multilingual support.
        // Change to "en-US" for English or "fr-FR" for French.
        CultureInfo.CurrentCulture = new CultureInfo("fr-FR");

        // Build the data model.
        ReportModel model = new()
        {
            Url = "https://example.com",
            LinkText = GetLocalizedString("LinkText")
        };

        // Create the LINQ Reporting template programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template and build the report.
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Model class used as the root data source for the report.
    public class ReportModel
    {
        public string Url { get; set; } = "";
        public string LinkText { get; set; } = "";
    }

    // Creates the template document containing a LINQ Reporting link tag.
    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Insert a paragraph with a link tag. The display text comes from the data model.
        builder.Writeln("Visit our site: <<link [model.Url] [model.LinkText]>>");

        doc.Save(path);
    }

    // Retrieves a localized string from the appropriate .resx file.
    private static string GetLocalizedString(string key)
    {
        string baseName = "Strings";
        string cultureSuffix = CultureInfo.CurrentCulture.TwoLetterISOLanguageName == "fr" ? ".fr" : "";
        string resxFile = $"{baseName}{cultureSuffix}.resx";

        if (!File.Exists(resxFile))
            return key; // Fallback if resource file is missing.

        try
        {
            XDocument doc = XDocument.Load(resxFile);
            XElement dataElement = doc.Root?.Element("data") ??
                                   doc.Root?.Elements("data")
                                      .FirstOrDefault(e => (string)e.Attribute("name") == key);

            // If the simple lookup above fails, search all <data> elements.
            if (dataElement == null)
            {
                foreach (XElement element in doc.Root?.Elements("data") ?? [])
                {
                    if ((string)element.Attribute("name") == key)
                    {
                        dataElement = element;
                        break;
                    }
                }
            }

            if (dataElement != null)
            {
                XElement valueElement = dataElement.Element("value");
                if (valueElement != null)
                    return valueElement.Value;
            }
        }
        catch
        {
            // Ignore parsing errors and fall back to the key.
        }

        return key; // Fallback if key is not found.
    }

    // Writes simple .resx files for English (default) and French cultures.
    private static void CreateResourceFiles()
    {
        // Default (English) resources.
        string defaultResx = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
  <data name=""LinkText"" xml:space=""preserve"">
    <value>Click here</value>
  </data>
</root>";
        File.WriteAllText("Strings.resx", defaultResx);

        // French resources.
        string frenchResx = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
  <data name=""LinkText"" xml:space=""preserve"">
    <value>Cliquez ici</value>
  </data>
</root>";
        File.WriteAllText("Strings.fr.resx", frenchResx);
    }
}
