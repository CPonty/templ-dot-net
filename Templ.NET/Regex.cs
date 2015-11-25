using System.Text.RegularExpressions;

namespace TemplNET
{
    /// <summary>
    /// Represents a regular expression for a placeholder, given a supplied prefix
    /// </summary>
    /// <example>
    /// <code>
    /// var rxp = new TemplRegex("txt");
    ///
    /// var str = rxp.Text("Title"); // = "{txt:Title}"
    ///
    /// var pat = rxp.Pattern; // = "\{txt:([^\}]+)\}"
    /// </code>
    /// </example>
    public class TemplRegex
    {
        private const RegexOptions Options = RegexOptions.Singleline | RegexOptions.Compiled;

        public string Prefix;
        public Regex Pattern;

        /// <summary>
        /// Create a regular expression for a placeholder, given the supplied prefix
        /// </summary>
        public TemplRegex(string prefix)
        {
            Prefix = prefix;
            Pattern = BuildPattern(prefix);
        }

        /// <summary>
        /// Generates placeholder for the given body: string "{prefix:body}"
        /// </summary>
        /// <param name="body"></param>
        public string Text(string body)
        {
            return $"{TemplConst.MatchOpen}{Prefix}{TemplConst.FieldSep}{body}{TemplConst.MatchClose}";
        }
        private static Regex BuildPattern(string prefix)
        {
            return new Regex(
                       Regex.Escape($"{TemplConst.MatchOpen}{prefix}{TemplConst.FieldSep}") + @"([^" 
                     + Regex.Escape(TemplConst.MatchClose) + "]+)" 
                     + Regex.Escape(TemplConst.MatchClose), Options);
        }
    }
}