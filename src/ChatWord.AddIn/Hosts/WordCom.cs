using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ChatWord.AddIn.Hosts
{
    internal static class WordCom
    {
        private const BindingFlags GetFlags = BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public;
        private const BindingFlags SetFlags = BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public;
        private const BindingFlags CallFlags = BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public;
        private static readonly CultureInfo ComCulture = CultureInfo.GetCultureInfo("en-US");

        internal static object Get(object target, string name, params object[] args)
        {
            if (target == null) { throw new ArgumentNullException(nameof(target)); }
            return target.GetType().InvokeMember(name, GetFlags, null, target, args, ComCulture);
        }

        internal static object Call(object target, string name, params object[] args)
        {
            if (target == null) { throw new ArgumentNullException(nameof(target)); }
            return target.GetType().InvokeMember(name, CallFlags, null, target, args, ComCulture);
        }

        internal static void Set(object target, string name, params object[] args)
        {
            if (target == null) { throw new ArgumentNullException(nameof(target)); }
            target.GetType().InvokeMember(name, SetFlags, null, target, args, ComCulture);
        }

        internal static bool TryGet(object target, string name, out object value, params object[] args)
        {
            value = null;
            if (target == null) { return false; }
            try { value = Get(target, name, args); return true; }
            catch { return false; }
        }

        internal static bool TryCall(object target, string name, out object value, params object[] args)
        {
            value = null;
            if (target == null) { return false; }
            try { value = Call(target, name, args); return true; }
            catch { return false; }
        }

        internal static string String(object target, string name, string fallback = "")
        {
            return TryGet(target, name, out var value) && value != null
                ? Convert.ToString(value, CultureInfo.InvariantCulture)
                : fallback;
        }

        internal static string ReadText(object range)
        {
            var text = Get(range, "Text");
            if (text is string value) { return value; }
            if (text == null && Convert.ToInt32(Get(range, "Start")) == Convert.ToInt32(Get(range, "End")))
            { return string.Empty; }
            throw new WordToolException("READBACK_FAILED", "宿主未返回有效的 Range.Text 文本。");
        }

        internal static int Int(object target, string name, int fallback = 0)
        {
            try { return TryGet(target, name, out var value) && value != null ? Convert.ToInt32(value) : fallback; }
            catch { return fallback; }
        }

        internal static bool Bool(object target, string name, bool fallback = false)
        {
            try { return TryGet(target, name, out var value) && value != null ? Convert.ToBoolean(value) : fallback; }
            catch { return fallback; }
        }

        internal static void Release(object target)
        {
            try { if (target != null && Marshal.IsComObject(target)) { Marshal.ReleaseComObject(target); } }
            catch { }
        }

        internal static string VisibleText(object range, bool includeHidden = false)
        {
            var text = String(range, "Text", string.Empty) ?? string.Empty;
            text = text.Replace("\a", string.Empty).Replace("\r", "\n");
            if (!includeHidden && Bool(range, "Hidden", false)) { return string.Empty; }
            return text.TrimEnd('\n');
        }

        internal static object Item(object collection, object index)
        {
            if (collection == null) { return null; }
            try { return Call(collection, "Item", index); }
            catch { return Get(collection, "Item", index); }
        }
    }
}
