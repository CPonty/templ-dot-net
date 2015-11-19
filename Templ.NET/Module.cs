using System;
using System.Collections.Generic;
using System.Linq;
using Novacode;

namespace TemplNET
{
    public abstract class TemplModule<T> : TemplModule where T : TemplMatchPara, new()
    {
        public uint MinFields = 1;
        public uint MaxFields = 1;
        public new Func<T, DocX, object, T> CustomHandler = (m, doc, model) => m;

        public TemplModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        { }

        /// <summary>
        /// Add more placeholder prefixes to match against.
        /// </summary>
        /// <param name="prefixes"></param>
        public TemplModule<T> WithPrefixes(IEnumerable<string> prefixes)
        {
            prefixes.ToList().ForEach(s => AddPrefix(s));
            return this;
        }
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
        public void BuildFromScope(IEnumerable<T> scope)
        {
            // Note how we are constantly "committing" the changes using ToList().
            // This ensures order is preserved (e.g. all "finds" happen before all "handler"s)
            scope.Select( m => CheckFieldCount(m) ).ToList()
                 .Select( m => Handler(m)).ToList()
                 .Where(  m =>!m.RemoveExpired()).ToList()
                 .Select( m => CustomHandler(m, Doc, Model)).ToList()
                 .ForEach(m => m.RemoveExpired());
        }
        public override void Build()
        {
            BuildFromScope(Regexes.SelectMany(rxp => FindAll(rxp)));
        }
        /// <summary>
        /// Container-specific handler.
        /// Modify the underlying container; mark expired to delete
        /// </summary>
        /// <param name="m"></param>
        public abstract T Handler(T m);

        /// <summary>
        /// Container-specific finder.
        /// Get all instances from the document.
        /// </summary>
        /// <param name="rxp"></param>
        public abstract IEnumerable<T> FindAll(TemplRegex rxp);
    }

    public abstract class TemplModule
    {
        public string Name;
        protected TemplBuilder DocBuilder;
        public DocX Doc => DocBuilder.Doc;
        public object Model => DocBuilder.Model;
        public ISet<TemplRegex> Regexes = new HashSet<TemplRegex>();
        public Func<object, TemplBuilder, object> CustomHandler;

        public TemplModule(TemplBuilder docBuilder, string name, string prefix)
        {
            AddPrefix(prefix);
            DocBuilder = docBuilder;
            Name = name;
        }

        protected TemplModule AddPrefix(string prefix)
        {
            if (prefix.Contains(TemplConst.FieldSep))
            {
                throw new FormatException($"Templ: Module \"{Name}\": prefix \"{prefix}\" cannot contain the field separator '{TemplConst.FieldSep}'");
            }
            Regexes.Add(new TemplRegex(prefix));
            return this;
        }

        public abstract void Build();
    }
}