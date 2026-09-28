using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Attempt to set compatibility mode to Word2010 if the API supports it.
        // This uses reflection to avoid compile‑time errors on older Aspose.Words versions.
        var compatibilityOptions = doc.CompatibilityOptions;
        var modeProperty = compatibilityOptions.GetType().GetProperty("CompatibilityMode");
        if (modeProperty != null && modeProperty.CanWrite)
        {
            // The enum name for Word2010 compatibility is "Word2010".
            var enumValue = Enum.Parse(modeProperty.PropertyType, "Word2010");
            modeProperty.SetValue(compatibilityOptions, enumValue);
        }

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a multilevel list.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Level 0 – decimal numbers.
        ListLevel level0 = list.ListLevels[0];
        level0.NumberFormat = "%1.";
        level0.NumberStyle = NumberStyle.Arabic;
        level0.NumberPosition = 0;
        level0.Alignment = ListLevelAlignment.Left;
        level0.TextPosition = 30;

        // Level 1 – lower‑case letters.
        ListLevel level1 = list.ListLevels[1];
        level1.NumberFormat = "%2.";
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.NumberPosition = 30;
        level1.Alignment = ListLevelAlignment.Left;
        level1.TextPosition = 60;

        // Level 2 – bullet.
        ListLevel level2 = list.ListLevels[2];
        level2.NumberFormat = "•";
        level2.NumberStyle = NumberStyle.Bullet;
        level2.NumberPosition = 60;
        level2.Alignment = ListLevelAlignment.Left;
        level2.TextPosition = 90;

        // Add items to the list.
        builder.ListFormat.List = list;

        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");

        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Third level bullet item 1");
        builder.Writeln("Third level bullet item 2");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string outputPath = "ComplexList.docx";
        doc.Save(outputPath);

        // Verify the file was saved and can be reopened.
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
            // Document successfully loaded; no further action required.
        }
    }
}
