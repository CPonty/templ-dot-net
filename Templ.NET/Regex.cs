﻿using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

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
            return $"{{{Prefix}:{name}}}";
        }
        public static Regex BuildPattern(string prefix)
        {
            return new Regex(@"\{" + Regex.Escape(prefix) + @":([^\}]+)\}", Options);
        }
    }
}