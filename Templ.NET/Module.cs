using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Novacode;

namespace TemplNET
{
    /// <summary>
    /// Represents a code module for the document build process 
    /// </summary>
    /// Modules Find and Handle content (Matches) in the document.
    /// Building a full document template essentially involves executing the modules in a specified sequence.
    /// <seealso cref="TemplMatch"/>
    /// <seealso cref="Templ.DefaultModules"/>
    public abstract class TemplModule
    {
        /// <summary>
        /// Module instance name. Used as an identifier in the debugger.
        /// </summary>
        public string Name;
        /// <summary>
        /// Flag set if any Match instances were found during the Module's build process
        /// </summary>
        public bool Used = false;
        /// <summary>
        /// The set of placeholder regexes used to find Matches in the document
        /// </summary>
        public ISet<TemplRegex> Regexes = new HashSet<TemplRegex>();
        public IEnumerable<string> Prefixes => Regexes.Select(rxp => rxp.Prefix);
        /// <summary>
        /// Named metadata values. Included in the Module meta-report when building in debug mode.
        /// </summary>
        public TemplModuleStatistics Statistics = new TemplModuleStatistics();
        /// <summary>
        /// An optional second handler function. Assignable per module instance at runtime. 
        /// </summary>
        public Func<object, TemplDoc, object> CustomHandler;

        /// <summary>
        /// Creates a code module for the document build process
        /// </summary>
        /// <param name="name">Module name</param>
        /// <param name="prefix">Placeholder prefix when searching for matches</param>
        public TemplModule(string name, string prefix)
        {
            AddPrefix(prefix);
            Name = name;
        }

        /// <summary>
        /// Given a prefix, add a placeholder to <see cref="Regexes"/>. These are used to find Matches in the document.
        /// </summary>
        protected TemplModule AddPrefix(string prefix)
        {
            if (prefix.Contains(TemplConst.FieldSep))
            {
                throw new FormatException($"Templ: Module \"{Name}\": prefix \"{prefix}\" cannot contain the field separator '{TemplConst.FieldSep}'");
            }
            Regexes.Add(new TemplRegex(prefix));
            return this;
        }

        /// <summary>
        /// Applies all changes to the document.
        /// <see cref="Templ.Build"/> automatically executes it for all <see cref="Templ.ActiveModules"/>.
        /// </summary>
        public abstract void Build(DocX doc, object model);
    }

    /// <summary>
    /// Represents a code module for the document build process.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// Modules Find and Handle content (Matches) in the document.
    /// Building a full document template essentially involves executing the modules in a specified sequence.
    /// 
    /// Each concrete Module class is implemented for a specific Match type.
    /// 
    public abstract class TemplModule<T> : TemplModule where T : TemplMatchPara, new()
    {
        /// <summary>
        /// The minimum # of fields required for a valid placeholder Match. "{prefix:a:b:c}" contains 3 matches.
        /// </summary>
        /// <seealso cref="FindAll"/>
        public uint MinFields = 1;
        /// <summary>
        /// The maximum # of fields expected for a valid placeholder Match. "{prefix:a:b:c}" contains 3 matches.
        /// </summary>
        /// <seealso cref="FindAll"/>
        public uint MaxFields = 1;
        /// <summary>
        /// An optional second handler function. Assignable per module instance at runtime. 
        /// </summary>
        public new Func<DocX, object, T, T> CustomHandler = (doc, model, m) => m;

        /// <summary>
        /// Creates a code module for the document build process
        /// </summary>
        /// <param name="name">Module name</param>
        /// <param name="prefix">Placeholder prefix when searching for matches</param>
        public TemplModule(string name, string prefix)
            : base(name, prefix)
        { }

        /// <summary>
        /// Given a collection of prefixes, add placeholders to <see cref="TemplModule.Regexes"/>.
        /// </summary>
        /// These are used when searching for Matches in the document.
        public TemplModule<T> WithPrefixes(IEnumerable<string> prefixes)
        {
            prefixes.ToList().ForEach(s => AddPrefix(s));
            return this;
        }

        private bool RemoveExpired(T m)
        {
            if (m.RemoveExpired())
            {
                Statistics.removals++;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Verifies the number of fields in the supplied Match's placeholder is within the min/max expected for this Module.
        /// <para/> Throws exception if problems are found.
        /// </summary>
        /// <param name="m"></param>
        private T CheckFieldCount(T m)
        {
            var l = m.Fields.Length;
            if (l < MinFields)
            {
                throw new FormatException($"Templ: Module \"{GetType()}\" found a placeholder \"{m.Placeholder}\" with too few \"{TemplConst.FieldSep}\"-separated fields ({l}, minimum is {MinFields})");
            }
            if (l > MaxFields)
            {
                throw new FormatException($"Templ: Module \"{GetType()}\" found a placeholder \"{m.Placeholder}\" with too many \"{TemplConst.FieldSep}\"-separated fields ({l}; maxmimum is {MaxFields})");
            }
            return m;
        }
        /// <summary>
        /// Handle and/or remove a collection of Matches from the document
        /// </summary>
        public void BuildFromScope(DocX doc, object model, IEnumerable<T> scope)
        {
            var watch = Stopwatch.StartNew();
            // Mark module instance as "used" if any matches are being processed
            Used = (Used || scope.Count() > 0);
            watch.Stop();
            Statistics.matches += scope.Count();
            // Note how we are constantly "committing" the changes using ToList().
            // This ensures order is preserved (e.g. all "finds" happen before all "handler"s)
            scope.Select( m => CheckFieldCount(m) ).ToList()
                 .Select( m => Handler(doc, model, m)).ToList()
                 .Where(  m =>!RemoveExpired(m)).ToList()
                 .Select( m => CustomHandler(doc, model, m)).ToList()
                 .ForEach(m => RemoveExpired(m));
            Statistics.millis = watch.ElapsedMilliseconds;
        }
        /// <summary>
        /// Find and build all content from the document
        /// </summary>
        public override void Build(DocX doc, object model)
        {
            BuildFromScope(doc, model, Regexes.SelectMany(rxp => FindAll(doc, rxp)));
        }

        /// <summary>
        /// Module-specific Match handler.
        /// Implementations should modify the Matched content, or mark expired to delete
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="model"></param>
        /// <param name="m"></param>
        public abstract T Handler(DocX doc, object model, T m);

        /// <summary>
        /// Module-specific finder.
        /// Implementations should retrieve all regex-matching content from the document.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="rxp"></param>
        public abstract IEnumerable<T> FindAll(DocX doc, TemplRegex rxp);
    }

    public class TemplModuleStatistics
    {
        public long millis = 0;
        public int matches = 0;
        public int removals = 0;
        public Dictionary<string, string> statistics = new Dictionary<string, string>();
    }
}