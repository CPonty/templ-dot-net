using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace Templ
{
    /// <summary>
    /// 
    /// </summary>
    public class TemplRegex
    {
        private const RegexOptions Options = RegexOptions.Singleline | RegexOptions.Compiled;

        /// <summary>
        /// 
        /// </summary>
        public string Prefix;
        /// <summary>
        /// 
        /// </summary>
        public Regex Pattern;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="prefix"></param>
        public TemplRegex(string prefix)
        {
            Prefix = prefix;
            Pattern = BuildPattern(prefix);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public string Text(string name)
        {
            return $"{{{Prefix}:{name}}}";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns></returns>
        public static Regex BuildPattern(string prefix)
        {
            return new Regex(@"\{" + Regex.Escape(prefix) + @":([^\}]+)\}", Options);
        }
    }
}