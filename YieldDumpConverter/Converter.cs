using System.Text.RegularExpressions;

namespace YieldDumpConverter
{
    class Converter
    {
        /// <summary>
        /// Reformats a block of yield dump text so it looks nicer when copied into an Excel spreadsheet.
        /// </summary>
        /// <param name="inString">Raw yield dump text</param>
        /// <returns>Yield dump text formatted for a spreadsheet</returns>
        public static string Convert(string inString)
        {
            string convertedString = inString;

            // Change "...Date: 20150731" to "...Date: 07/31/2015".
            convertedString = Regex.Replace(
                convertedString,
                "^(?<label>\\w+Date): (?<year>\\d{4})(?<month>\\d{2})(?<dayOfMonth>\\d{2})",
                match => match.Groups["label"].Value + ": " + match.Groups["month"].Value + "/" + match.Groups["dayOfMonth"].Value + "/" + match.Groups["year"].Value,
                RegexOptions.Multiline | RegexOptions.IgnoreCase);

            // Change "\t20150731," to "\t07/31/2015,".
            convertedString = Regex.Replace(
                convertedString,
                "^(\t|( +))(?<year>\\d{4})(?<month>\\d{2})(?<dayOfMonth>\\d{2}),",
                match =>"\t" + match.Groups["month"].Value + "/" + match.Groups["dayOfMonth"].Value + "/" + match.Groups["year"].Value + ",",
                RegexOptions.Multiline | RegexOptions.IgnoreCase);

            // Change all occurrences of ": " to a tab character.
            convertedString = convertedString.Replace(": ", "\t");

            // Change all occurrences of ", " to a tab character.
            convertedString = convertedString.Replace(", ", "\t");

            // Other minor touch-up for a spreadsheet
            convertedString = convertedString.Replace("Cash Flows (Date", "Cash Flows\tDate");
            convertedString = convertedString.Replace("PresentValue):", "PresentValue");

            return convertedString;
        }

    }
}
