// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Word;

public class ListItemTextGetter_es_ES
{
    // 1–19 cardinal forms (index 0 = uno, index 18 = diecinueve)
    private static readonly string[] OneThroughNineteen =
    [
        "uno",
        "dos",
        "tres",
        "cuatro",
        "cinco",
        "seis",
        "siete",
        "ocho",
        "nueve",
        "diez",
        "once",
        "doce",
        "trece",
        "catorce",
        "quince",
        "dieciséis",
        "diecisiete",
        "dieciocho",
        "diecinueve",
    ];

    // 20–29 have fused forms; 21–29 are "veinti" + ones
    private static readonly string[] Veinti =
    [
        "veintiuno",
        "veintidós",
        "veintitrés",
        "veinticuatro",
        "veinticinco",
        "veintiséis",
        "veintisiete",
        "veintiocho",
        "veintinueve",
    ];

    // Tens 30–90 (index 0 = treinta … index 6 = noventa)
    private static readonly string[] Tens =
    [
        "treinta",
        "cuarenta",
        "cincuenta",
        "sesenta",
        "setenta",
        "ochenta",
        "noventa",
    ];

    // Hundreds (index 0 = ciento … index 8 = novecientos)
    // 100 alone = "cien"; in compounds 101–199 = "ciento"
    private static readonly string[] HundredsPrefix =
    [
        "ciento",
        "doscientos",
        "trescientos",
        "cuatrocientos",
        "quinientos",
        "seiscientos",
        "setecientos",
        "ochocientos",
        "novecientos",
    ];

    // 1–10 ordinal forms (index 0 = primero … index 9 = décimo)
    private static readonly string[] OrdinalOneThroughTen =
    [
        "primero",
        "segundo",
        "tercero",
        "cuarto",
        "quinto",
        "sexto",
        "séptimo",
        "octavo",
        "noveno",
        "décimo",
    ];

    // 11–19 ordinal forms
    private static readonly string[] OrdinalElevenThroughNineteen =
    [
        "undécimo",
        "duodécimo",
        "decimotercero",
        "decimocuarto",
        "decimoquinto",
        "decimosexto",
        "decimoséptimo",
        "decimoctavo",
        "decimonoveno",
    ];

    // Tens ordinals 20–90 (index 0 = vigésimo … index 6 = nonagésimo)
    private static readonly string[] OrdinalTens =
    [
        "vigésimo",
        "trigésimo",
        "cuadragésimo",
        "quincuagésimo",
        "sexagésimo",
        "septuagésimo",
        "octogésimo",
        "nonagésimo",
    ];

    private static string BuildCardinal(int n)
    {
        if (n <= 0 || n > 19999)
            return n.ToString();

        var result = "";

        // Thousands
        var t1 = n / 1000;
        var t2 = n % 1000;
        if (t1 == 1)
        {
            result += "mil";
        }
        else if (t1 >= 2)
        {
            result += BuildCardinal(t1) + " mil";
        }

        if (t1 >= 1 && t2 == 0)
            return result;
        if (t1 >= 1)
            result += " ";

        // Hundreds
        var h1 = (n % 1000) / 100;
        var h2 = n % 100;
        if (h1 >= 1)
        {
            if (h1 == 1 && h2 == 0)
                result += "cien"; // exactly 100
            else
                result += HundredsPrefix[h1 - 1];
        }

        if (h1 >= 1 && h2 == 0)
            return result;
        if (h1 >= 1)
            result += " ";

        // Last two digits
        var z = n % 100;
        if (z <= 19)
        {
            result += OneThroughNineteen[z - 1];
        }
        else if (z <= 29)
        {
            if (z == 20)
                result += "veinte";
            else
                result += Veinti[z - 21];
        }
        else
        {
            var x = z / 10; // 3–9 → index 0–6 in Tens
            var r = z % 10;
            result += Tens[x - 3];
            if (r >= 1)
                result += " y " + OneThroughNineteen[r - 1];
        }

        return result;
    }

    public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
    {
        if (numFmt == "cardinalText")
        {
            if (levelNumber <= 0 || levelNumber > 19999)
                return levelNumber.ToString();

            var result = BuildCardinal(levelNumber);
            return result[0..1].ToUpper() + result[1..];
        }

        if (numFmt == "ordinalText")
        {
            if (levelNumber <= 0 || levelNumber > 19999)
                return levelNumber.ToString();

            var z = levelNumber % 100;
            string? ordinalSuffix;

            if (z >= 1 && z <= 10)
                ordinalSuffix = OrdinalOneThroughTen[z - 1];
            else if (z >= 11 && z <= 19)
                ordinalSuffix = OrdinalElevenThroughNineteen[z - 11];
            else
            {
                // 20–99: tens ordinal + optional units ordinal
                var x = z / 10; // 2–9 → index 0–6 in OrdinalTens
                var r = z % 10;
                ordinalSuffix = z == 0 ? null : OrdinalTens[x - 2];
                if (r >= 1 && r <= 10)
                    ordinalSuffix = (ordinalSuffix is null ? "" : ordinalSuffix + " ") + OrdinalOneThroughTen[r - 1];
            }

            // Build the thousands/hundreds prefix (cardinal) and attach the ordinal
            var t1 = levelNumber / 1000;
            var h1 = (levelNumber % 1000) / 100;
            var h2 = levelNumber % 100;
            var result = "";

            if (t1 == 1)
                result += "milésimo";
            else if (t1 >= 2)
                result += BuildCardinal(t1) + " milésimo";

            if (t1 >= 1 && levelNumber % 1000 == 0)
            {
                var r2 = result;
                return r2[0..1].ToUpper() + r2[1..];
            }
            if (t1 >= 1)
                result += " ";

            if (h1 >= 1)
            {
                var hundredsOrdinal = h1 switch
                {
                    1 => "centésimo",
                    2 => "ducentésimo",
                    3 => "tricentésimo",
                    4 => "cuadringentésimo",
                    5 => "quingentésimo",
                    6 => "sexcentésimo",
                    7 => "septingentésimo",
                    8 => "octingentésimo",
                    9 => "noningentésimo",
                    _ => "",
                };
                result += hundredsOrdinal;
                if (h2 == 0)
                {
                    var r2 = result;
                    return r2[0..1].ToUpper() + r2[1..];
                }
                result += " ";
            }

            if (ordinalSuffix is not null)
                result += ordinalSuffix;
            else
                result = result.TrimEnd();

            return result[0..1].ToUpper() + result[1..];
        }

        return null;
    }
}
