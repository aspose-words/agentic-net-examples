using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Paths for the XML configuration and the output document.
        const string xmlPath = "fontSubstitutions.xml";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample XML configuration file that defines font substitutes.
        // -----------------------------------------------------------------
        var xmlDoc = new XDocument(
            new XElement("FontSubstitutions",
                new XElement("Substitution",
                    new XAttribute("original", "MissingFont"),
                    new XElement("Substitute", "Arial"),
                    new XElement("Substitute", "Times New Roman")
                ),
                new XElement("Substitution",
                    new XAttribute("original", "AnotherMissingFont"),
                    new XElement("Substitute", "Calibri")
                )
            )
        );
        xmlDoc.Save(xmlPath);

        // -----------------------------------------------------------------
        // Load the XML and configure Aspose.Words FontSettings.
        // -----------------------------------------------------------------
        var loadedXml = XDocument.Load(xmlPath);
        var fontSettings = new FontSettings();

        // The Aspose.Words API for adding substitutes has changed across versions.
        // To stay compatible we use reflection to call the appropriate member.
        var substitutionSettings = fontSettings.SubstitutionSettings;
        PropertyInfo tableProp = substitutionSettings.GetType().GetProperty("Table");
        MethodInfo addSubstitutesMethod = substitutionSettings.GetType().GetMethod("AddSubstitutes");

        foreach (var substitution in loadedXml.Root.Elements("Substitution"))
        {
            string originalFont = substitution.Attribute("original")?.Value;
            if (string.IsNullOrWhiteSpace(originalFont))
                continue;

            string[] substitutes = substitution.Elements("Substitute")
                                                .Select(e => e.Value.Trim())
                                                .Where(v => !string.IsNullOrWhiteSpace(v))
                                                .ToArray();

            if (substitutes.Length == 0)
                continue;

            // Preferred API: substitutionSettings.Table.AddSubstitutes(...)
            if (tableProp != null)
            {
                var tableInstance = tableProp.GetValue(substitutionSettings);
                MethodInfo addMethod = tableInstance.GetType().GetMethod("AddSubstitutes", new[] { typeof(string), typeof(string[]) });
                addMethod?.Invoke(tableInstance, new object[] { originalFont, substitutes });
            }
            // Fallback API: substitutionSettings.AddSubstitutes(...)
            else if (addSubstitutesMethod != null)
            {
                addSubstitutesMethod.Invoke(substitutionSettings, new object[] { originalFont, substitutes });
            }
        }

        // -----------------------------------------------------------------
        // Build a document that uses a font that is unlikely to be installed.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font and should be substituted.");

        // Apply the custom FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        bool fileExists = File.Exists(outputPath);
        Console.WriteLine($"Document saved: {fileExists}");
    }
}
