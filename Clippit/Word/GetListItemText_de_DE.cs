// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Word;

public class ListItemTextGetter_de_DE
{
    // 1–19 cardinal forms (index 0 = ein, index 18 = neunzehn)
    private static readonly string[] OneThroughNineteen =
    [
        "ein",
        "zwei",
        "drei",
        "vier",
        "fünf",
        "sechs",
        "sieben",
        "acht",
        "neun",
        "zehn",
        "elf",
        "zwölf",
        "dreizehn",
        "vierzehn",
        "fünfzehn",
        "sechzehn",
        "siebzehn",
        "achtzehn",
        "neunzehn",
    ];

    // Tens (index 0 = zwanzig … index 7 = neunzig)
    private static readonly string[] Tens =
    [
        "zwanzig",
        "dreißig",
        "vierzig",
        "fünfzig",
        "sechzig",
        "siebzig",
        "achtzig",
        "neunzig",
    ];

    // Hundreds (index 0 = hundert … index 8 = neunhundert)
    private static readonly string[] Hundreds =
    [
        "hundert",
        "zweihundert",
        "dreihundert",
        "vierhundert",
        "fünfhundert",
        "sechshundert",
        "siebenhundert",
        "achthundert",
        "neunhundert",
    ];

    // 1–19 ordinal adjective forms (weak declension, masculine nominative)
    // index 0 = erste … index 18 = neunzehnte
    private static readonly string[] OrdinalOneThroughNineteen =
    [
        "erste",
        "zweite",
        "dritte",
        "vierte",
        "fünfte",
        "sechste",
        "siebte",
        "achte",
        "neunte",
        "zehnte",
        "elfte",
        "zwölfte",
        "dreizehnte",
        "vierzehnte",
        "fünfzehnte",
        "sechzehnte",
        "siebzehnte",
        "achtzehnte",
        "neunzehnte",
    ];

    /// <summary>Builds the German cardinal word form for 1–19999.</summary>
    private static string CardinalCore(int n)
    {
        if (n <= 0 || n > 19999)
            return n.ToString();

        var result = "";

        // Thousands
        var t1 = n / 1000;
        if (t1 == 1)
            result += "tausend";
        else if (t1 >= 2)
            result += OneThroughNineteen[t1 - 1] + "tausend";

        // Hundreds
        var h1 = (n % 1000) / 100;
        if (h1 >= 1)
            result += Hundreds[h1 - 1];

        // Last two digits (0–99)
        var z = n % 100;
        if (z == 0)
            return result;

        if (z <= 19)
        {
            result += OneThroughNineteen[z - 1];
        }
        else
        {
            var x = z / 10; // tens index (2 = zwanzig, …, 9 = neunzig)
            var r = z % 10;
            if (r == 0)
            {
                result += Tens[x - 2];
            }
            else
            {
                // German: units BEFORE tens joined with "und"
                result += OneThroughNineteen[r - 1] + "und" + Tens[x - 2];
            }
        }

        return result;
    }

    public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
    {
        if (numFmt == "cardinalText")
        {
            if (levelNumber <= 0 || levelNumber > 19999)
                return levelNumber.ToString();

            var result = CardinalCore(levelNumber);
            return result[0..1].ToUpper() + result[1..];
        }

        if (numFmt == "ordinalText")
        {
            if (levelNumber <= 0 || levelNumber > 19999)
                return levelNumber.ToString();

            var result = "";

            // Thousands prefix (always cardinal)
            var t1 = levelNumber / 1000;
            var t2 = levelNumber % 1000;
            if (t1 == 1)
                result += "tausend";
            else if (t1 >= 2)
                result += OneThroughNineteen[t1 - 1] + "tausend";

            if (t1 >= 1 && t2 == 0)
            {
                result += "ste";
                return result[0..1].ToUpper() + result[1..];
            }

            // Hundreds prefix (always cardinal)
            var h1 = (levelNumber % 1000) / 100;
            var h2 = levelNumber % 100;
            if (h1 >= 1)
                result += Hundreds[h1 - 1];

            if (h1 >= 1 && h2 == 0)
            {
                result += "ste";
                return result[0..1].ToUpper() + result[1..];
            }

            // Last two digits
            var z = levelNumber % 100;
            if (z <= 19)
            {
                result += OrdinalOneThroughNineteen[z - 1];
            }
            else
            {
                var x = z / 10;
                var r = z % 10;
                if (r == 0)
                {
                    // 20th, 30th, …, 90th: cardinal + "ste"
                    result += Tens[x - 2] + "ste";
                }
                else
                {
                    // 21st, 22nd, …: [units]und[tens]ste
                    result += OneThroughNineteen[r - 1] + "und" + Tens[x - 2] + "ste";
                }
            }

            return result[0..1].ToUpper() + result[1..];
        }

        return null;
    }
}
