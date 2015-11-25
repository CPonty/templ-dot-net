using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace TemplNET
{
    /// <summary>
    /// Represents an object in the document, located by matching a Regex placeholder string.
    /// </summary>
    /// <seealso cref="TemplMatchPara"/>
    /// 'Match' is abstract. Matching of specific objects (e.g. pictures, tables) is implemented by subclasses.
    public abstract class TemplMatch
    {
        /// <summary>
        /// Flag set by handler modules. Indicates the matched object will be automatically removed.
        /// </summary>
        public bool Expired = false;
        /// <summary>
        /// Flag set by handler modules. Indicates the matched object has been removed.
        /// </summary>
        public bool Removed = false;
        /// <summary>
        /// Flag set by handler modules. Indicates the placeholder string has been removed from the document.
        /// </summary>
        protected bool RemovedPlaceholder = false;
        /// <summary>
        /// The contents of the placeholder string located in the document. "{prefix:Body}"
        /// </summary>
        public string Body;
        /// <summary>
        /// Generates an array of ':'-separated strings from the placeholder body
        /// </summary>
        public string[] Fields => Body.Split(TemplConst.FieldSep);
        /// <summary>
        /// The Regex placeholder used to find the Match
        /// </summary>
        public TemplRegex Regex;
        /// <summary>
        /// Retrieves the raw placeholder string found in the document
        /// </summary>
        public string Placeholder => Regex.Text(Body);
        /// <summary>
        /// Retrieves the base Regex object used to find the Match
        /// </summary>
        public Regex Pattern => Regex.Pattern;

        /* if someone tries to implement single-result search: .DefaultIfEmpty(new T()).First(); */

        /// <summary>
        /// Find matches in paragraphs, based on a Regex placeholder pattern.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgx"></param>
        /// <param name="paragraphs"></param>
        public static IEnumerable<T> Find<T>(TemplRegex rgx, IEnumerable<Paragraph> paragraphs) where T : TemplMatchPara, new()
        {
            return paragraphs.SelectMany(p => Find<T>(rgx, p));
        }
        /// <summary>
        /// Find matches in paragraphs, based on a Regex placeholder pattern. Limit n per paragraph
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgx"></param>
        /// <param name="paragraphs"></param>
        /// <param name="n"></param>
        public static IEnumerable<T> Find<T>(TemplRegex rgx, IEnumerable<Paragraph> paragraphs, int n) where T : TemplMatchPara, new()
        {
            return paragraphs.SelectMany(p => Find<T>(rgx, p).Take(n));
        }
        /// <summary>
        /// Find matches in a paragraph, based on a Regex placeholder pattern.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgx"></param>
        /// <param name="p"></param>
        public static IEnumerable<T> Find<T>(TemplRegex rgx, Paragraph p) where T : TemplMatchPara, new()
        {
            var P = p;
            var t = P.Text;
            return Find<T>(rgx, p.Text).Select(m =>
            {
                m.Paragraph = p;
                return m;
            });
        }
        /// <summary>
        /// Find matches in a string, based on a Regex placeholder pattern.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgx"></param>
        /// <param name="s"></param>
        public static IEnumerable<T> Find<T>(TemplRegex rgx, string s) where T : TemplMatch, new()
        {
            return rgx.Pattern.Matches(s).Cast<Match>().Where(m => m.Success).Select(m => new T()
            {
                Regex = rgx,
                Body = m.Groups[1].Value,
            });
        }
    }

    /// <summary>
    /// Represents an object in the document, located by matching a Regex placeholder string in the scope of a Paragraph.
    /// </summary>
    /// <seealso cref="TemplMatch"/>
    /// 'MatchPara' is still abstract. Matching of specific objects (e.g. pictures, tables) is implemented by subclassing.
    public abstract class TemplMatchPara : TemplMatch
    {
        /// <summary>
        /// The paragraph this Match was located in
        /// </summary>
        public Paragraph Paragraph;

        /// <summary>
        /// Insert string at the given index of the paragraph.
        /// </summary>
        /// <param name="idx"></param>
        /// <param name="text"></param>
        private void InsertText(int idx, string text)
        {
            // Simple. Text at index
            Paragraph.InsertText(idx, text);
        }

        /// <summary>
        /// Replace the placeholder with a string in the paragraph.
        /// </summary>
        /// <param name="text"></param>
        public T ToText<T>(string text) where T : TemplMatchPara, new()
        {
            // Insert Text at each match, then remove placeholder
            Paragraph.ReplaceText(Placeholder, text);
            RemovedPlaceholder = true;
            return (T)this;
        }
        /// <summary>
        /// Insert a string before the paragraph. Creates a new paragraph.
        /// </summary>
        public Paragraph AddBefore<T>(string text) where T : TemplMatchPara, new() => Paragraph.InsertParagraphBeforeSelf(text);
        /// <summary>
        /// Insert a string after the paragraph. Creates a new paragraph.
        /// </summary>
        public Paragraph AddAfter<T>(string text) where T : TemplMatchPara, new() => Paragraph.InsertParagraphAfterSelf(text);

        /// <summary>
        /// Insert a picture at the beginning of the paragraph. The graphic must already be loaded into the same document.
        /// </summary>
        private void InsertPicture(TemplGraphic graphic, int width)
        {
            // Simple. Picture at index(es). Ensure you load picture first.
            Paragraph.FindAll(Placeholder).ForEach(idx =>
            {
                if (width>0)
                {
                    Paragraph.InsertPicture(graphic.Picture((uint)width), idx + Placeholder.Length);
                }
                else
                {
                    Paragraph.InsertPicture(graphic.Picture(), idx + Placeholder.Length);
                }
                Paragraph.Alignment = (graphic.Alignment ?? Paragraph.Alignment);
            });
        }
        /// <summary>
        /// Replace the placeholder with a picture in the paragraph. The graphic must already be loaded into the same document.
        /// </summary>
        public T ToPicture<T>(TemplGraphic graphic, int width) where T : TemplMatchPara, new()
        {
            // Use InsertPicture at each match, then remove placeholder
            if (Removed || RemovedPlaceholder || Expired)
            {
                return (T)this;
            }
            InsertPicture(graphic,width);
            RemovePlaceholder();
            return (T)this;
        }

        /// <summary>
        /// Replace the placeholder with a collection of pictures. The graphics must already be loaded into the same document.
        /// </summary>
        public T ToPictures<T>(ICollection<TemplGraphic> graphics, int width) where T : TemplMatchPara, new()
        {
            return ToPictures<T>(graphics.ToArray(), width);
        }

        /// <summary>
        /// Replace the placeholder with an array of pictures. The graphics must already be loaded into the same document.
        /// </summary>
        public T ToPictures<T>(TemplGraphic[] graphics, int width) where T : TemplMatchPara, new()
        {
            if (Removed || RemovedPlaceholder || Expired)
            {
                return (T)this;
            }
            // Use InsertPicture for each graphic, then remove placeholder.
            // Reverse order as pictures get inserted immediately after the placeholder, 
            //  so each subsequent picture inserted will appear before those already inserted.
            foreach (var graphic in graphics.Reverse())
            {
                InsertPicture(graphic, width);
            }
            RemovePlaceholder();
            return (T)this;
        }

        /// <summary>
        /// Remove the matched content from the document
        /// </summary>
        /// By default, only the placeholder is removed.
        /// Subclasses should override Remove() to comprehensively remove the relevant matched object.
        public virtual void Remove()
        {
            if (Removed)
            {
                return;
            }
            RemovePlaceholder();
            Removed = true;
        }

        /// <summary>
        /// Remove the matched placeholder string from the document
        /// </summary>
        public virtual void RemovePlaceholder()
        {
            if (RemovedPlaceholder)
            {
                return;
            }
            Paragraph?.ReplaceText(Placeholder, string.Empty);
            RemovedPlaceholder = true;
        }

        /// <summary>
        /// Remove if the matched object is marked as expired
        /// </summary>
        public bool RemoveExpired()
        {
            if (Expired)
            {
                Remove();
            }
            return Removed;
        }
    }
}