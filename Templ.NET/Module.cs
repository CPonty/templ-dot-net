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
        public new Func<DocX, object, T, T> CustomHandler = (doc, model, m) => m;

        public TemplModule(string name, string prefix)
            : base(name, prefix)
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
        public void BuildFromScope(DocX doc, object model, IEnumerable<T> scope)
        {
            // Note how we are constantly "committing" the changes using ToList().
            // This ensures order is preserved (e.g. all "finds" happen before all "handler"s)
            scope.Select( m => CheckFieldCount(m) ).ToList()
                 .Select( m => Handler(doc, model, m)).ToList()
                 .Where(  m =>!m.RemoveExpired()).ToList()
                 .Select( m => CustomHandler(doc, model, m)).ToList()
                 .ForEach(m => m.RemoveExpired());
        }
        public override void Build(DocX doc, object model)
        {
            BuildFromScope(doc, model, Regexes.SelectMany(rxp => FindAll(doc, rxp)));
        }
        /// <summary>
        /// Container-specific handler.
        /// Modify the underlying container; mark expired to delete
        /// </summary>
        /// <param name="m"></param>
        /// <param name="doc"></param>
        /// <param name="model"></param>
        public abstract T Handler(DocX doc, object model, T m);

        /// <summary>
        /// Container-specific finder.
        /// Get all instances from the document.
        /// </summary>
        /// <param name="rxp"></param>
        /// <param name="doc"></param>
        public abstract IEnumerable<T> FindAll(DocX doc, TemplRegex rxp);
    }

    public abstract class TemplModule
    {
        public string Name;
        public ISet<TemplRegex> Regexes = new HashSet<TemplRegex>();
        public Func<object, TemplBuilder, object> CustomHandler;

        public TemplModule(string name, string prefix)
        {
            AddPrefix(prefix);
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

        public abstract void Build(DocX doc, object model);
    }
}