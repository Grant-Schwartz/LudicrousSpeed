using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace LudicrousSpeed.OutlookAddIn.Services
{
    /// <summary>
    /// What one workbook file turned out to contain.
    /// </summary>
    internal sealed class WorkbookScan
    {
        /// <summary>
        /// Number of LS.LIVE/WS.LIVE formulas found. This is the number that
        /// matters: every one of them is a cell that shows #NAME? on a machine
        /// without the add-in.
        /// </summary>
        public int LiveCellCount { get; set; }

        /// <summary>Sheets those cells sit on, in workbook order.</summary>
        public List<string> Sheets { get; } = new List<string>();

        /// <summary>
        /// True when the workbook still uses the retired WS.LIVE name, which
        /// is already broken even for the sender -- worth saying out loud.
        /// </summary>
        public bool UsesLegacyName { get; set; }

        /// <summary>
        /// The hidden conversion-metadata sheet is present. On its own this is
        /// harmless -- it survives Restore Native on purpose -- so it never
        /// triggers a warning, it only gets mentioned in one that already fired.
        /// </summary>
        public bool HasConversionMetadata { get; set; }

        /// <summary>
        /// Set when the file could not be read far enough to make a claim.
        /// Callers must treat this as "don't know", never as "clean".
        /// </summary>
        public string? Inconclusive { get; set; }

        public bool NeedsAttention => LiveCellCount > 0;
    }

    /// <summary>
    /// Detects LudicrousSpeed live cells in a saved workbook file.
    ///
    /// WHY THIS EXISTS: Convert to Live replaces a native Excel data table
    /// with =LS.LIVE("Sheet!Cell") formulas backed by an RTD server inside the
    /// add-in. That is fine on the analyst's machine and useless anywhere
    /// else: a recipient without LudicrousSpeed installed has no such
    /// function, so every converted cell resolves to #NAME? and the
    /// sensitivity tables in the model read as broken. Restore Native puts the
    /// real tables back, and this scanner is what reminds someone to run it
    /// before the file leaves the building.
    ///
    /// WHY IT READS THE FILE RATHER THAN ASKING EXCEL: the attachment is a
    /// saved copy on disk. It may not be open, Excel may not be running, and
    /// the workbook may have been converted in a session that has since
    /// ended. The .xlsx is a zip of XML, so the evidence is right there.
    ///
    /// WHY IT SCANS ONLY WORKSHEET PARTS: formula text lives inline in a
    /// worksheet's &lt;f&gt; elements, while cell *string values* are pooled
    /// into xl/sharedStrings.xml. Skipping the shared string table means a
    /// cell whose text happens to read "LS.LIVE" cannot be mistaken for a
    /// converted cell.
    /// </summary>
    internal static class LiveCellScanner
    {
        /// <summary>
        /// Extensions worth opening. Everything else is skipped without
        /// touching the file, which is what keeps an ordinary attachment of a
        /// PDF or a photo free.
        /// </summary>
        private static readonly HashSet<string> WorkbookExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam",
            };

        private const string LiveFunction = "LS.LIVE";
        private const string LegacyLiveFunction = "WS.LIVE";
        private const string MetadataSheetName = "_LudicrousSpeed_DataTables";

        /// <summary>
        /// Ceiling on decompressed bytes examined per workbook. A model big
        /// enough to blow through this is also one nobody wants to wait on
        /// mid-send, and without a ceiling a hostile or corrupt zip could
        /// expand without bound. Hitting it makes the scan inconclusive rather
        /// than clean.
        /// </summary>
        private const long ByteBudget = 512L * 1024 * 1024;

        public static bool LooksLikeWorkbook(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }

            try
            {
                return WorkbookExtensions.Contains(Path.GetExtension(fileName));
            }
            catch (ArgumentException)
            {
                // Invalid path characters in an attachment name; not ours.
                return false;
            }
        }

        public static WorkbookScan Scan(string path)
        {
            try
            {
                using (var zip = ZipFile.OpenRead(path))
                {
                    return ScanArchive(zip, IsBinaryWorkbook(path));
                }
            }
            catch (InvalidDataException)
            {
                // Not a zip at all: an old .xls renamed, or a truncated file.
                return new WorkbookScan { Inconclusive = "not a readable OpenXML workbook" };
            }
            catch (Exception ex)
            {
                return new WorkbookScan { Inconclusive = ex.Message };
            }
        }

        private static bool IsBinaryWorkbook(string path)
        {
            return string.Equals(Path.GetExtension(path), ".xlsb", StringComparison.OrdinalIgnoreCase);
        }

        private static WorkbookScan ScanArchive(ZipArchive zip, bool binary)
        {
            var scan = new WorkbookScan();
            var budget = ByteBudget;

            // .xlsb keeps its formulas in a tokenized binary stream, but the
            // name of an add-in function still lands there as a UTF-16LE
            // string, so the same substring search works on a different
            // encoding. Both encodings are searched either way: it costs one
            // extra comparison per byte and removes a whole class of
            // "detection quietly stopped working" bug.
            var markers = new[]
            {
                Utf8(LiveFunction), Utf16(LiveFunction),
                Utf8(LegacyLiveFunction), Utf16(LegacyLiveFunction),
            };

            var sheetNames = binary ? new SheetMap() : ReadSheetMap(zip);
            scan.HasConversionMetadata = sheetNames.Names.Any(
                name => string.Equals(name, MetadataSheetName, StringComparison.OrdinalIgnoreCase));

            foreach (var entry in zip.Entries)
            {
                if (!IsWorksheetPart(entry.FullName, binary))
                {
                    continue;
                }

                if (budget <= 0)
                {
                    break;
                }

                int[] counts;
                using (var stream = entry.Open())
                {
                    counts = ByteSearch.Count(stream, markers, ref budget);
                }

                var live = counts[0] + counts[1];
                var legacy = counts[2] + counts[3];
                if (live + legacy == 0)
                {
                    continue;
                }

                scan.LiveCellCount += live + legacy;
                scan.UsesLegacyName |= legacy > 0;
                scan.Sheets.Add(DescribeSheet(sheetNames, entry.FullName));
            }

            // Set after the loop rather than inside it, so a budget exhausted
            // partway through the final worksheet is reported too.
            if (budget <= 0)
            {
                scan.Inconclusive = "workbook is larger than this check will read";
            }

            if (binary && scan.LiveCellCount > 0)
            {
                // No sheet name map for a binary workbook, so the metadata
                // sheet has to be spotted the same way its cells were.
                scan.HasConversionMetadata = ContainsMetadataSheetName(zip);
            }

            return scan;
        }

        private static bool IsWorksheetPart(string entryPath, bool binary)
        {
            if (!entryPath.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // xl/worksheets/_rels/... and the drawing/comment parts alongside
            // the sheets hold no formulas.
            if (entryPath.IndexOf("/_rels/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            return entryPath.EndsWith(binary ? ".bin" : ".xml", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsMetadataSheetName(ZipArchive zip)
        {
            var workbook = zip.Entries.FirstOrDefault(e =>
                e.FullName.Equals("xl/workbook.bin", StringComparison.OrdinalIgnoreCase));
            if (workbook == null)
            {
                return false;
            }

            var budget = 8L * 1024 * 1024;
            using (var stream = workbook.Open())
            {
                var counts = ByteSearch.Count(
                    stream,
                    new[] { Utf8(MetadataSheetName), Utf16(MetadataSheetName) },
                    ref budget);
                return counts[0] + counts[1] > 0;
            }
        }

        private static string DescribeSheet(SheetMap map, string entryPath)
        {
            if (map.ByEntry.TryGetValue(entryPath, out var name) && !string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            // Better than nothing when the relationship graph is unusual:
            // "sheet4.xml" at least tells someone which sheet to look at.
            return Path.GetFileName(entryPath);
        }

        private sealed class SheetMap
        {
            public Dictionary<string, string> ByEntry { get; } =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            public List<string> Names { get; } = new List<string>();
        }

        /// <summary>
        /// Maps each worksheet part back to the tab name a user would
        /// recognise, by joining xl/workbook.xml's sheet list to
        /// xl/_rels/workbook.xml.rels. Purely cosmetic -- naming the sheets in
        /// the warning is the difference between "go find it" and "go here" --
        /// so every failure here degrades to part file names rather than
        /// abandoning the scan.
        /// </summary>
        private static SheetMap ReadSheetMap(ZipArchive zip)
        {
            var map = new SheetMap();
            try
            {
                var workbook = FindEntry(zip, "xl/workbook.xml");
                if (workbook == null)
                {
                    return map;
                }

                var relationships = ReadRelationships(zip);

                XDocument document;
                using (var stream = workbook.Open())
                {
                    document = XDocument.Load(stream);
                }

                var sheets = document.Root
                    ?.Elements().FirstOrDefault(e => e.Name.LocalName == "sheets")
                    ?.Elements().Where(e => e.Name.LocalName == "sheet")
                    ?? Enumerable.Empty<XElement>();

                foreach (var sheet in sheets)
                {
                    var name = (string?)sheet.Attribute("name") ?? "";
                    map.Names.Add(name);

                    // The r:id attribute is namespace-qualified; match on local
                    // name so an unexpected prefix or namespace revision does
                    // not silently break the lookup.
                    var relationshipId = sheet.Attributes()
                        .FirstOrDefault(a => a.Name.LocalName == "id")?.Value;
                    if (string.IsNullOrEmpty(relationshipId))
                    {
                        continue;
                    }

                    if (relationships.TryGetValue(relationshipId!, out var target))
                    {
                        map.ByEntry[target] = name;
                    }
                }
            }
            catch (Exception)
            {
                // Cosmetic only -- see the summary above.
            }

            return map;
        }

        private static Dictionary<string, string> ReadRelationships(ZipArchive zip)
        {
            var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
            var entry = FindEntry(zip, "xl/_rels/workbook.xml.rels");
            if (entry == null)
            {
                return relationships;
            }

            XDocument document;
            using (var stream = entry.Open())
            {
                document = XDocument.Load(stream);
            }

            foreach (var element in document.Root?.Elements() ?? Enumerable.Empty<XElement>())
            {
                if (element.Name.LocalName != "Relationship")
                {
                    continue;
                }

                var id = (string?)element.Attribute("Id");
                var target = (string?)element.Attribute("Target");
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target))
                {
                    continue;
                }

                relationships[id!] = ResolveTarget(target!);
            }

            return relationships;
        }

        /// <summary>
        /// Turns a relationship target into a zip entry path. Targets are
        /// relative to the part that declared them -- xl/workbook.xml here --
        /// except when they are absolute, which Excel does not emit but the
        /// format allows.
        /// </summary>
        private static string ResolveTarget(string target)
        {
            var path = target.Replace('\\', '/');
            if (path.StartsWith("/", StringComparison.Ordinal))
            {
                return path.TrimStart('/');
            }

            var segments = new List<string>(("xl/" + path).Split('/'));
            var resolved = new List<string>();
            foreach (var segment in segments)
            {
                if (segment == "." || segment.Length == 0)
                {
                    continue;
                }

                if (segment == "..")
                {
                    if (resolved.Count > 0)
                    {
                        resolved.RemoveAt(resolved.Count - 1);
                    }

                    continue;
                }

                resolved.Add(segment);
            }

            return string.Join("/", resolved);
        }

        private static ZipArchiveEntry? FindEntry(ZipArchive zip, string path)
        {
            return zip.Entries.FirstOrDefault(
                e => e.FullName.Equals(path, StringComparison.OrdinalIgnoreCase));
        }

        private static byte[] Utf8(string value)
        {
            return Encoding.UTF8.GetBytes(value);
        }

        private static byte[] Utf16(string value)
        {
            return Encoding.Unicode.GetBytes(value);
        }
    }
}
