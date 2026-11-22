using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;

namespace ClosedXML.Excel.Drawings
{
    internal class XLPictures : IXLPictures, IEnumerable<XLPicture>
    {
        private readonly List<XLPicture> _pictures = new List<XLPicture>();
        private readonly XLWorksheet _worksheet;

        public XLPictures(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
            Deleted = new HashSet<String>();
        }

        public int Count
        {
            [DebuggerStepThrough]
            get { return _pictures.Count; }
        }

        internal ICollection<String> Deleted { get; private set; }

        public IXLPicture Add(Stream stream)
        {
            var picture = new XLPicture(_worksheet, stream);
            _pictures.Add(picture);
            picture.Name = GetNextPictureName();
            return picture;
        }

        public IXLPicture Add(Stream stream, string name)
        {
            var picture = Add(stream);
            picture.Name = name;
            return picture;
        }

        public IXLPicture Add(Stream stream, XLPictureFormat format)
        {
            var picture = new XLPicture(_worksheet, stream, format);
            _pictures.Add(picture);
            picture.Name = GetNextPictureName();
            return picture;
        }

        public IXLPicture Add(Stream stream, XLPictureFormat format, string name)
        {
            var picture = Add(stream, format);
            picture.Name = name;
            return picture;
        }

        public IXLPicture Add(string imageFile)
        {
            using (var fs = File.OpenRead(imageFile))
            {
                var picture = new XLPicture(_worksheet, fs);
                _pictures.Add(picture);
                picture.Name = GetNextPictureName();
                return picture;
            }
        }

        public IXLPicture Add(string imageFile, string name)
        {
            var picture = Add(imageFile);
            picture.Name = name;
            return picture;
        }

        public bool Contains(string pictureName)
        {
            return _pictures.Any(p => string.Equals(p.Name, pictureName, StringComparison.OrdinalIgnoreCase));
        }

        public void Delete(IXLPicture picture)
        {
            Delete(picture.Name);
        }

        public void Delete(string pictureName)
        {
            var picturesToDelete = _pictures
                .Where(picture => picture.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (!picturesToDelete.Any())
                throw new ArgumentOutOfRangeException(nameof(pictureName), $"Picture {pictureName} was not found.");

            foreach (var picture in picturesToDelete)
            {
                if (!string.IsNullOrEmpty(picture.RelId))
                    Deleted.Add(picture.RelId);

                _pictures.Remove(picture);
            }
        }

        IEnumerator<IXLPicture> IEnumerable<IXLPicture>.GetEnumerator()
        {
            return _pictures.Cast<IXLPicture>().GetEnumerator();
        }

        public IEnumerator<XLPicture> GetEnumerator()
        {
            return ((IEnumerable<XLPicture>)_pictures).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLPicture Picture(string pictureName)
        {
            if (TryGetPicture(pictureName, out IXLPicture? p))
                return p;

            throw new ArgumentOutOfRangeException(nameof(pictureName), $"Picture {pictureName} was not found.");
        }

        public bool TryGetPicture(string pictureName, [NotNullWhen(true)] out IXLPicture? picture)
        {
            var match = _pictures.FirstOrDefault(p => p.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase));
            if (match is not null)
            {
                picture = match;
                return true;
            }
            picture = null;
            return false;
        }

        /// <summary>
        /// Add a loaded picture. The picture uses the <paramref name="imageStream"/> without a copy
        /// and owns it.
        /// </summary>
        internal XLPicture AddOwned(MemoryStream imageStream, string name, int id)
        {
            var picture = XLPicture.FromOwnedStream(_worksheet, imageStream);
            _pictures.Add(picture);
            picture.Name = GetNextPictureName();

            // Excel can save a picture with an empty name. Keep the generated name for it, as Excel does.
            if (!String.IsNullOrEmpty(name))
                picture.SetName(name);

            picture.Id = id;
            return picture;
        }

        /// <summary>
        /// Give a new generated name to each of the <paramref name="pictures"/> that has the name of a different picture.
        /// </summary>
        internal void RenameDuplicates(IEnumerable<XLPicture> pictures)
        {
            foreach (var picture in pictures)
            {
                var hasDuplicate = _pictures.Any(p => p != picture && p.Name.Equals(picture.Name, StringComparison.OrdinalIgnoreCase));
                if (hasDuplicate)
                    picture.SetName(GetNextPictureName());
            }
        }

        private String GetNextPictureName()
        {
            var pictureNumber = this.Count;
            while (_pictures.Any(p => p.Name.Equals($"Picture {pictureNumber}", StringComparison.OrdinalIgnoreCase)))
            {
                pictureNumber++;
            }
            return $"Picture {pictureNumber}";
        }
    }
}
