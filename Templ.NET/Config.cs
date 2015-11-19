using System.Collections.Generic;
using System.Linq;

namespace TemplNET
{
    public static class TemplConfig
    {
        public const bool Debug = false;
        public const string MatchOpen = "{";
        public const string MatchClose = "}";
        public const char FieldSep = ':';
        public const uint MaxMatchesPerScope = 999;
    }
}