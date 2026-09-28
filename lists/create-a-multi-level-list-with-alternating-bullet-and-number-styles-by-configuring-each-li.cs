using System;
using System.Collections.Generic;

public class Program
{
    public static void Main()
    {
        var items = new List<ListItem>
        {
            new ListItem(0, "First top‑level item"),
            new ListItem(1, "First sub‑item"),
            new ListItem(2, "First sub‑sub‑item"),
            new ListItem(1, "Second sub‑item"),
            new ListItem(0, "Second top‑level item"),
            new ListItem(1, "Another sub‑item")
        };

        PrintMultiLevelList(items);
    }

    private static void PrintMultiLevelList(List<ListItem> items)
    {
        // Determine the deepest level to size helper arrays.
        int maxLevel = 0;
        foreach (var item in items)
            if (item.Level > maxLevel) maxLevel = item.Level;

        // Counters for numbered levels.
        int[] counters = new int[maxLevel + 1];

        // Define style per level: true = numbered, false = bullet.
        // Alternating: even levels numbered, odd levels bullet.
        bool[] isNumbered = new bool[maxLevel + 1];
        for (int i = 0; i <= maxLevel; i++)
            isNumbered[i] = i % 2 == 0;

        foreach (var item in items)
        {
            // Reset deeper level counters when moving up the hierarchy.
            for (int lvl = item.Level + 1; lvl <= maxLevel; lvl++)
                counters[lvl] = 0;

            string indent = new string(' ', item.Level * 4);
            if (isNumbered[item.Level])
            {
                counters[item.Level]++;
                Console.WriteLine($"{indent}{counters[item.Level]}. {item.Text}");
            }
            else
            {
                Console.WriteLine($"{indent}- {item.Text}");
            }
        }
    }

    private class ListItem
    {
        public int Level { get; }
        public string Text { get; }

        public ListItem(int level, string text)
        {
            Level = level;
            Text = text;
        }
    }
}
