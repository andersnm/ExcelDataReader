using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsBiffHyperlink : XlsBiffRecord
    {
        private static readonly Guid StdLink = new Guid("79EAC9D0-BAF9-11CE-8C82-00AA004BA90B");
        private static readonly Guid UrlMoniker = new Guid(0x79EAC9E0, 0xBAF9, 0x11CE, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B);
        private static readonly Guid FileMoniker = new Guid(0x00000303, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
        private static readonly Guid CompositeMoniker = new Guid(0x00000309, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
        private static readonly Guid AntiMoniker = new Guid(0x00000305, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
        private static readonly Guid ItemMoniker = new Guid(0x00000304, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);

        internal XlsBiffHyperlink(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
            // 8 bytes cell range
            FirstRow = ReadUInt16(0x0);
            LastRow = ReadUInt16(0x2);
            FirstColumn = ReadUInt16(0x4);
            LastColumn = ReadUInt16(0x6);

            // 16 byte guid, must be "StdLink"
            var linkGuid = new Guid(ReadArray(0x8, 16));

            if (linkGuid != StdLink)
            {
                return;
            }

            // [MS-OSHARED] 2.3.7.1 Hyperlink Object
            var streamVersion = ReadUInt32(0x18); // always 2
            var optionFlags = (OptionFlags)ReadUInt32(0x1C);

            var hyperlinkOffset = 0x20;

            if ((optionFlags & OptionFlags.HasDisplayName) != 0)
            {
                // Variable length display name
                DisplayName = ReadHyperlinkString(hyperlinkOffset, out var bytesRead);
                hyperlinkOffset += bytesRead;
            }

            if ((optionFlags & OptionFlags.HasFrameName) != 0)
            {
                // Variable length target frame name
                var targetFrameName = ReadHyperlinkString(hyperlinkOffset, out var bytesRead);
                hyperlinkOffset += bytesRead;
            }

            if ((optionFlags & OptionFlags.HasMoniker) != 0)
            {
                if ((optionFlags & OptionFlags.MonikerSavedAsStr) != 0)
                {
                    // Variable length moniker string
                    Moniker = ReadHyperlinkString(hyperlinkOffset, out var bytesRead);
                    hyperlinkOffset += bytesRead;
                }
                else
                {
                    // Variable length olemoniker
                    Moniker = ReadHyperLinkMoniker(hyperlinkOffset, out var bytesRead);
                    hyperlinkOffset += bytesRead;
                }
            }

            if ((optionFlags & OptionFlags.HasLocationStr) != 0)
            {
                // variable length location
                Location = ReadHyperlinkString(hyperlinkOffset, out var bytesRead);
                hyperlinkOffset += bytesRead;
            }

            if ((optionFlags & OptionFlags.HasGUID) != 0)
            {
                // fixed length GUID
            }

            if ((optionFlags & OptionFlags.HasCreationTime) != 0)
            {
                // fixed length creation time
            }

            // throw new Exception("us");
        }

        internal enum OptionFlags : uint
        {
            HasMoniker = 0x1,
            IsAbsolute = 0x2,
            SiteGaveDisplayName = 0x4,
            HasLocationStr = 0x8,
            HasDisplayName = 0x10,
            HasGUID = 0x20,
            HasCreationTime = 0x40,
            HasFrameName = 0x80,
            MonikerSavedAsStr = 0x100,
            AbsFromGetdataRel = 0x200
        }

        public int FirstRow { get; }

        public int LastRow { get; }

        public int FirstColumn { get; }

        public int LastColumn { get; }

        public string DisplayName { get; }

        public string Moniker { get; }

        public string Location { get; }

        /// <summary>
        /// 2.3.7.9 HyperlinkString
        /// </summary>
        private string ReadHyperlinkString(int offset, out int bytesRead)
        {
            var characterCount = ReadInt32(offset);
            var bytes = ReadArray(offset + 4, characterCount * 2);
            bytesRead = 4 + characterCount * 2;
            return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
        }

        /// <summary>
        /// 2.3.7.2 HyperlinkMoniker
        /// </summary>
        private string ReadHyperLinkMoniker(int offset, out int bytesRead)
        {
            var monikerClsid = new Guid(ReadArray(offset, 16));
            if (monikerClsid == UrlMoniker)
            {
                var urlOffset = offset + 16;

                var dataSize = ReadUInt32(urlOffset);
                urlOffset += 4;

                // Count unicode characters until zero terminator
                var characterCount = 0;
                while (characterCount * 2 < dataSize)
                {
                    var c = ReadUInt16(urlOffset + characterCount * 2);
                    if (c == 0)
                        break;
                    characterCount++;
                }

                bytesRead = 16 + 4 + characterCount * 2;

                // NOTE: Skipping serialGUID, serialVersion, uriFlags present if there is exactly 24 bytes remaining
                return Encoding.Unicode.GetString(ReadArray(urlOffset, characterCount * 2));
            }
            else if (monikerClsid == FileMoniker)
            {
                var fileOffset = offset + 16;

                var anti = ReadUInt16(fileOffset);
                fileOffset += 2;

                var ansiLength = ReadInt32(fileOffset);
                fileOffset += 4;

                var ansiPath = ReadArray(fileOffset, ansiLength);
                fileOffset += ansiLength;

                var endServer = ReadUInt16(fileOffset);
                fileOffset += 2;

                var versionNumber = ReadUInt16(fileOffset);
                fileOffset += 2;

                // Skip reserved chunks
                fileOffset += 16 + 4;

                var unicodePathSize = ReadInt32(fileOffset);
                fileOffset += 4;

                if (unicodePathSize != 0)
                {
                    // Is a unicode path, read cbUnicodePathBytes, usKeyValue, and unicodePath
                    var unicodePathBytes = ReadInt32(fileOffset);
                    fileOffset += 4;

                    var keyValue = ReadUInt16(fileOffset);
                    fileOffset += 2;

                    var unicodePath = Encoding.Unicode.GetString(ReadArray(fileOffset, unicodePathBytes));
                    fileOffset += unicodePathBytes;

                    bytesRead = fileOffset - offset;
                    return unicodePath;
                }
                else
                {
                    // ansiPath is viable! No more data
                    bytesRead = fileOffset - offset;
                    
                    // NOTE/TODO: what encoding is "ANSI"? assume lower 8 bits of unicode aka iso8859-1??
                    throw new InvalidOperationException("ANSI URL");
                }
            }
            else if (monikerClsid == CompositeMoniker)
            {
                // UNTESTED
                var compositeOffset = offset + 16;

                var monikerCount = ReadUInt32(offset);
                compositeOffset += 4;

                for (var i = 0; i < monikerCount; i++)
                {
                    ReadHyperLinkMoniker(compositeOffset, out var compositeReadBytes);
                    compositeOffset += compositeReadBytes;
                }

                bytesRead = compositeOffset - offset;
                return null;
            }
            else if (monikerClsid == AntiMoniker)
            {
                // UNTESTED
                bytesRead = 16 + 4;
                return null;
            }
            else if (monikerClsid == ItemMoniker)
            {
                var itemOffset = offset + 16;

                var delimiterLength = ReadInt32(itemOffset);
                itemOffset += 4;
                itemOffset += delimiterLength;

                var itemLength = ReadInt32(itemOffset);
                itemOffset += 4;
                itemOffset += itemLength;

                bytesRead = itemOffset - offset;
                return null;
                throw new InvalidOperationException("ItemMoniker");
            }
            else
            {
                throw new InvalidOperationException("Unexpected hyperlink moniker CLSID " + monikerClsid.ToString());
            }
        }
    }
}
