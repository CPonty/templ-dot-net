using System.Text.RegularExpressions;

namespace TemplNET
{
    public class TemplRegex
    {
        private const RegexOptions Options = RegexOptions.Singleline | RegexOptions.Compiled;

        public string Prefix;
        public Regex Pattern;

        public TemplRegex(string prefix)
        {
            Prefix = prefix;
            Pattern = BuildPattern(prefix);
        }

        public string Text(string name)
        {
            return $"{TemplConst.MatchOpen}{Prefix}{TemplConst.FieldSep}{name}{TemplConst.MatchClose}";
        }
        public static Regex BuildPattern(string prefix)
        {
            return new Regex(
                       Regex.Escape($"{TemplConst.MatchOpen}{prefix}{TemplConst.FieldSep}") + @"([^" 
                     + Regex.Escape(TemplConst.MatchClose) + "]+)" 
                     + Regex.Escape(TemplConst.MatchClose), Options);
        }
    }
}