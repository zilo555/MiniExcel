using System.Collections;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace MiniExcelLibs;

/// <summary>
/// Provides experimental XLSX queries backed by the native MiniExcel Rust library.
/// </summary>
public static class MiniExcelRust
{
    private const int BatchSize = 64;

    /// <summary>
    /// Streams rows from an XLSX file through the native Rust query engine.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> Query(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1")
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));
        if (string.IsNullOrWhiteSpace(startCell))
            throw new ArgumentException("The start cell is required.", nameof(startCell));

        return QueryIterator(Path.GetFullPath(path), useHeaderRow, sheetName, startCell);
    }

    private static IEnumerable<IDictionary<string, object?>> QueryIterator(
        string path,
        bool useHeaderRow,
        string? sheetName,
        string startCell)
    {
        EnsureAbiVersion();

        using var nativePath = new Utf8String(path);
        using var nativeSheetName = new Utf8String(sheetName);
        using var nativeStartCell = new Utf8String(startCell);
        var result = NativeMethods.QueryOpen(
            nativePath.Pointer,
            useHeaderRow ? (byte)1 : (byte)0,
            nativeSheetName.Pointer,
            nativeStartCell.Pointer,
            out var rawHandle);
        if (result < 0)
            throw CreateNativeException(result);

        using var handle = new NativeQueryHandle(rawHandle);
        while (true)
        {
            result = NativeMethods.QueryNextBatch(handle, BatchSize, out var data, out var length);
            if (result == 0)
                yield break;
            if (result < 0)
                throw CreateNativeException(result);

            var byteLength = checked((int)length.ToUInt64());
            var frame = new byte[byteLength];
            Marshal.Copy(data, frame, 0, byteLength);
            foreach (var row in DecodeBatch(frame))
                yield return row;
        }
    }

    private static IEnumerable<IDictionary<string, object?>> DecodeBatch(byte[] frame)
    {
        var reader = new FrameReader(frame);
        var rowCount = reader.ReadLength();
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var cellCount = reader.ReadLength();
            IDictionary<string, object?> row = new Dictionary<string, object?>(cellCount, StringComparer.Ordinal);
            for (var cellIndex = 0; cellIndex < cellCount; cellIndex++)
                row.Add(reader.ReadString(), reader.ReadValue());
            yield return row;
        }

        reader.EnsureComplete();
    }

    private static void EnsureAbiVersion()
    {
        var version = NativeMethods.GetAbiVersion();
        if (version != 1)
            throw new NotSupportedException($"MiniExcel Rust ABI version {version} is not supported.");
    }

    private static Exception CreateNativeException(int result)
    {
        var data = NativeMethods.GetLastError(out var length);
        var byteLength = checked((int)length.ToUInt64());
        if (data == IntPtr.Zero || byteLength == 0)
            return new InvalidOperationException($"MiniExcel Rust query failed with native error {result}.");

        var bytes = new byte[byteLength];
        Marshal.Copy(data, bytes, 0, byteLength);
        return new InvalidOperationException(Encoding.UTF8.GetString(bytes));
    }

    private sealed class FrameReader(byte[] frame)
    {
        private int _offset;

        public int ReadLength()
        {
            var value = ReadUInt32();
            if (value > int.MaxValue)
                throw new InvalidDataException("The native MiniExcel frame contains an unsupported length.");
            return (int)value;
        }

        public string ReadString()
        {
            var length = ReadLength();
            EnsureAvailable(length);
            var value = Encoding.UTF8.GetString(frame, _offset, length);
            _offset += length;
            return value;
        }

        public object? ReadValue()
        {
            EnsureAvailable(1);
            return frame[_offset++] switch
            {
                0 => null,
                1 => ReadBoolean(),
                2 => Convert.ToDouble(ReadInt64(), CultureInfo.InvariantCulture),
                3 => BitConverter.Int64BitsToDouble(ReadInt64()),
                4 => ReadString(),
                5 => DateTime.ParseExact(ReadString(), "yyyy-MM-dd", CultureInfo.InvariantCulture),
                6 => TimeSpan.Parse(ReadString(), CultureInfo.InvariantCulture),
                7 => DateTime.Parse(ReadString(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
                8 => TimeSpan.FromMilliseconds(ReadInt64()),
                9 => ReadString(),
                var tag => throw new InvalidDataException($"The native MiniExcel frame contains unknown value tag {tag}.")
            };
        }

        public void EnsureComplete()
        {
            if (_offset != frame.Length)
                throw new InvalidDataException("The native MiniExcel frame contains trailing data.");
        }

        private bool ReadBoolean()
        {
            EnsureAvailable(1);
            return frame[_offset++] != 0;
        }

        private uint ReadUInt32()
        {
            EnsureAvailable(sizeof(uint));
            var value = (uint)(frame[_offset]
                | frame[_offset + 1] << 8
                | frame[_offset + 2] << 16
                | frame[_offset + 3] << 24);
            _offset += sizeof(uint);
            return value;
        }

        private long ReadInt64()
        {
            EnsureAvailable(sizeof(long));
            ulong value = 0;
            for (var index = 0; index < sizeof(long); index++)
                value |= (ulong)frame[_offset + index] << (index * 8);
            _offset += sizeof(long);
            return unchecked((long)value);
        }

        private void EnsureAvailable(int length)
        {
            if (length < 0 || _offset > frame.Length - length)
                throw new InvalidDataException("The native MiniExcel frame is truncated.");
        }
    }

    private sealed class Utf8String : IDisposable
    {
        public Utf8String(string? value)
        {
            if (value is null)
                return;

            var bytes = Encoding.UTF8.GetBytes(value);
            Pointer = Marshal.AllocHGlobal(bytes.Length + 1);
            Marshal.Copy(bytes, 0, Pointer, bytes.Length);
            Marshal.WriteByte(Pointer, bytes.Length, 0);
        }

        public IntPtr Pointer { get; private set; }

        public void Dispose()
        {
            if (Pointer == IntPtr.Zero)
                return;

            Marshal.FreeHGlobal(Pointer);
            Pointer = IntPtr.Zero;
        }
    }

    private sealed class NativeQueryHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public NativeQueryHandle() : base(true) { }

        public NativeQueryHandle(IntPtr value) : this()
        {
            SetHandle(value);
        }

        protected override bool ReleaseHandle()
        {
            NativeMethods.QueryClose(handle);
            return true;
        }
    }

    private static class NativeMethods
    {
        private const string LibraryName = "miniexcel_ffi";

        [DllImport(LibraryName, EntryPoint = "miniexcel_abi_version", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern uint GetAbiVersion();

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_open", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int QueryOpen(
            IntPtr path,
            byte useHeaderRow,
            IntPtr sheetName,
            IntPtr startCell,
            out IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_next_batch", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int QueryNextBatch(
            NativeQueryHandle handle,
            uint maxRows,
            out IntPtr data,
            out UIntPtr length);

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_close", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern void QueryClose(IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_last_error", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern IntPtr GetLastError(out UIntPtr length);
    }
}