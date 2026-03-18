using System;
using Aspose.Words;
using Aspose.Words.Lists;

class AdjustListLevelIndentation
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a list based on a predefined template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of the first three list levels.
        for (int i = 0; i < 3; i++)
        {
            ListLevel level = list.ListLevels[i];

            // In current Aspose.Words versions the ListLevel class does NOT expose an Indentation property.
            // Indentation is defined by the combination of NumberPosition (where the number is placed) and
            // TextPosition (where the list text starts). To achieve a 36‑point (0.5 inch) indent we set:
            //   NumberPosition = -36  -> number is placed 36 points to the left of the left margin.
            //   TextPosition   = 36  -> list text starts 36 points to the right of the left margin.
            level.NumberPosition = -36.0; // number left of the indent
            level.TextPosition = 36.0;    // text starts after the indent
        }

        // Use DocumentBuilder to add some list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        builder.Writeln("First level item");
        builder.ListFormat.ListIndent(); // go to second level
        builder.Writeln("Second level item");
        builder.ListFormat.ListIndent(); // go to third level
        builder.Writeln("Third level item");
        builder.ListFormat.ListOutdent(); // back to second level
        builder.Writeln("Another second level item");
        builder.ListFormat.ListOutdent(); // back to first level
        builder.Writeln("Another first level item");

        // Remove list formatting from the builder.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("AdjustedListIndentation.docx");
    }
}
