using System;
using System.IO;

namespace LudicrousSpeed.OutlookAddIn.Services
{
    /// <summary>
    /// Counts occurrences of short byte patterns in a forward-only stream.
    ///
    /// WHY NOT JUST ReadToEnd AND String.IndexOf: the streams here are
    /// decompressed worksheet parts, and a worksheet in a real model routinely
    /// decompresses to tens of megabytes. Materialising that as a string costs
    /// twice the bytes again in UTF-16 and puts it on the large object heap,
    /// inside Outlook's process, on every attachment. Scanning a fixed 64 KiB
    /// window costs the same wall time and a constant amount of memory.
    ///
    /// All patterns are matched in one pass, ASCII-case-insensitively, so the
    /// caller can look for a name in several encodings and casings at once.
    /// </summary>
    internal static class ByteSearch
    {
        private const int ChunkSize = 64 * 1024;

        /// <param name="budget">
        /// Decompressed bytes this call may read, decremented as it goes. Kept
        /// shared across an archive so one workbook cannot be made to cost
        /// unbounded time by a zip that expands enormously.
        /// </param>
        public static int[] Count(Stream stream, byte[][] patterns, ref long budget)
        {
            var counts = new int[patterns.Length];
            if (patterns.Length == 0)
            {
                return counts;
            }

            var lowered = new byte[patterns.Length][];
            var longest = 0;
            for (var i = 0; i < patterns.Length; i++)
            {
                lowered[i] = ToLowerAscii(patterns[i]);
                longest = Math.Max(longest, lowered[i].Length);
            }

            // Room for a full chunk plus the tail of the previous one, so a
            // pattern straddling a chunk boundary is still seen exactly once.
            var buffer = new byte[ChunkSize + longest];
            var carried = 0;

            while (budget > 0)
            {
                var wanted = (int)Math.Min(ChunkSize, budget);
                var read = stream.Read(buffer, carried, wanted);
                if (read <= 0)
                {
                    break;
                }

                budget -= read;
                var filled = carried + read;

                // Start positions past this point cannot hold a whole pattern
                // yet; they get carried forward instead of being scanned twice.
                var scanTo = Math.Max(0, filled - longest + 1);
                Scan(buffer, 0, scanTo, filled, lowered, counts);

                carried = filled - scanTo;
                Buffer.BlockCopy(buffer, scanTo, buffer, 0, carried);
            }

            // Whatever is left is shorter than the longest pattern, but not
            // necessarily shorter than every pattern.
            Scan(buffer, 0, carried, carried, lowered, counts);
            return counts;
        }

        private static void Scan(
            byte[] buffer,
            int from,
            int to,
            int filled,
            byte[][] lowered,
            int[] counts)
        {
            for (var i = from; i < to; i++)
            {
                for (var p = 0; p < lowered.Length; p++)
                {
                    if (MatchesAt(buffer, i, filled, lowered[p]))
                    {
                        counts[p]++;
                    }
                }
            }
        }

        private static bool MatchesAt(byte[] buffer, int start, int filled, byte[] lowerPattern)
        {
            if (start + lowerPattern.Length > filled)
            {
                return false;
            }

            for (var k = 0; k < lowerPattern.Length; k++)
            {
                if (LowerAscii(buffer[start + k]) != lowerPattern[k])
                {
                    return false;
                }
            }

            return true;
        }

        private static byte[] ToLowerAscii(byte[] pattern)
        {
            var lowered = new byte[pattern.Length];
            for (var i = 0; i < pattern.Length; i++)
            {
                lowered[i] = LowerAscii(pattern[i]);
            }

            return lowered;
        }

        /// <summary>
        /// ASCII-only fold. That is exactly the right scope: it covers the
        /// case Excel actually varies (function names round-tripped through
        /// different writers) and leaves the zero bytes of a UTF-16LE pattern
        /// alone, so the same comparison serves both encodings.
        /// </summary>
        private static byte LowerAscii(byte value)
        {
            return value >= (byte)'A' && value <= (byte)'Z'
                ? (byte)(value + 32)
                : value;
        }
    }
}
