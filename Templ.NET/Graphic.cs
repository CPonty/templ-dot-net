using System;
using System.IO;
using Novacode;

namespace TemplNET
{
    /// <summary>
    /// Used to load/store images and generate formatted pictures for the document
    /// </summary>
    /// <example>
    /// <code>
    /// using Novacode;
    ///
    /// var graphic1 = new TemplGraphic("C:\logo.png");
    /// var graphic2 = new TemplGraphic("C:\photo.jpg", Alignment.Left, 0.5);
    /// 
    /// // Include images in the model
    /// Templ.Load("C:\template.docx")
    ///      .Build( new { Title = "Hello World!", logo = graphic1, pic = graphic2 })
    ///      .SaveAs("C:\output.docx");
    /// </code>
    /// </example>
    public class TemplGraphic
    {
        /// <summary>
        /// Underlying DocX image instance
        /// <para/>Not populated until document build time, as it requires a reference to a DocX document
        /// </summary>
        protected Novacode.Image Image;

        /// <summary>
        /// Image data as memory stream
        /// </summary>
        public MemoryStream Stream;
        /// <summary>
        /// Image data as bytes
        /// </summary>
        public byte[] Bytes => Stream.ToArray();

        /// <summary>
        /// Resizing factor when inserting into the document. Maintains aspect ratio
        /// </summary>
        public double Scalar;
        /// <summary>
        /// Alignment when inserting into the document (e.g. Left, Right, Centre). No change in alignment if null
        /// </summary>
        public Alignment? Alignment;
        /// <summary>
        /// String description of the image. Has no automatic function; useful for storing caption text in the model.
        /// </summary>
        public string Description;

        private bool Loaded = false;

        /// <summary>
        /// New graphic from data stream
        /// </summary>
        public TemplGraphic(Stream data, Alignment? align = null, double scalar = 1.0, string description = "")
        {
            Stream = new MemoryStream();
            data.CopyTo(Stream);
            Scalar = scalar;
            Alignment = align;
            Description = description;
        }

        /// <summary>
        /// New graphic from file data
        /// </summary>
        public TemplGraphic(byte[] data, Alignment? align = null, double scalar = 1.0, string description = "")
            : this(new MemoryStream(data), align, scalar, description)
        { }

        /// <summary>
        /// New graphic from filename
        /// </summary>
        public TemplGraphic(string file, Alignment? align = null, double scalar = 1.0, string description = "")
            : this(File.ReadAllBytes(file), align, scalar, description)
        { }

        /// <summary>
        /// Load image into document.
        /// Required before producing a Picture for the document.
        /// </summary>
        public TemplGraphic Load(DocX doc)
        {
            if (!Loaded)
            {
                Image = doc.AddImage(Stream);
                Loaded = true;
            }
            return this;
        }

        /// <summary>
        /// Generate a Picture instance, which can be inserted into a Document
        /// </summary>
        /// <param name="scalar">The scaling factor (0.5 is 50%)</param>
        public Picture Picture(double scalar)
        {
            if (!Loaded)
            {
                throw new InvalidDataException("Templ: Tried to generate a Picture from a Graphic before the Graphic was loaded into a Document");
            }
            var pic = Image.CreatePicture();
            pic.Width = (int)(pic.Width * (scalar));
            pic.Height = (int)(pic.Height * (scalar));
            return pic;
        }

        /// <summary>
        /// Generate a Picture instance, which can be inserted into a Document
        /// </summary>
        public Picture Picture()
        {
            return Picture(Scalar);
        }

        /// <summary>
        /// Generate a Picture instance, which can be inserted into a Document
        /// </summary>
        /// <param name="width">The width in pixels</param>
        public Picture Picture(uint width)
        {
            var pic = Image.CreatePicture();
            var scaling = ((float)width / pic.Width);
            pic.Height = (int)(pic.Height * scaling);
            pic.Width = (int)(pic.Width * scaling);
            return pic;
        }
    }
}