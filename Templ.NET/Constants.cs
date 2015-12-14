using System.Collections.Generic;
using System.Linq;

namespace TemplNET
{
    /// <summary>
    /// Global constants for Templ.NET
    /// </summary>
    public static class TemplConst
    {
        public const bool Debug = false;
        public const string MatchOpen = "{";
        public const string MatchClose = "}";
        public const char FieldSep = ':';
        public const uint MaxMatchesPerScope = 999;
        public static class Prefix
        {
            public const string Section = "sec";
            public const string Picture = "pic";
            public const string List = "li";
            public const string Table = "tab";
            public const string Row = "row";
            public const string Cell = "cel";
            public const string Cells = "cells";
            public const string Hyperlink = "url";
            public const string Text = "txt";
            public const string Remove = "rm";
            public const string Contents = "toc";
            public const string Comment = "!";
            public const string Collection = "$";
        }
    }
}