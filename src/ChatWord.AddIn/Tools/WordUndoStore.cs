using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ChatWord.AddIn.Hosts;

namespace ChatWord.AddIn.Tools
{
    internal sealed class WordUndoStore
    {
        private sealed class Entry
        {
            internal WeakReference Document;
            internal IntPtr ComIdentity;
            internal int StoryType;
            internal int Start;
            internal int End;
            internal string Before;
            internal string After;
            internal bool Consumed;
        }

        private readonly Dictionary<string, Entry> _entries = new Dictionary<string, Entry>(StringComparer.Ordinal);

        internal void Add(string id, object document, WordTarget target, string before, string after)
        {
            if (string.IsNullOrWhiteSpace(id) || target == null) { return; }
            _entries[id] = new Entry
            {
                Document = new WeakReference(document), ComIdentity = Identity(document), StoryType = target.StoryType,
                Start = target.Start, End = target.Start + (after ?? string.Empty).Length,
                Before = before ?? string.Empty, After = after ?? string.Empty,
            };
        }

        internal bool TryUndo(string id, object document, out string message)
        {
            message = null;
            if (!_entries.TryGetValue(id ?? string.Empty, out var entry) || entry.Consumed)
            { message = "没有可用的文档快照撤销记录。"; return false; }
            if (document == null || (entry.ComIdentity != IntPtr.Zero
                ? Identity(document) != entry.ComIdentity : !ReferenceEquals(entry.Document.Target, document)))
            { message = "当前文档不是生成快照时的文档实例。"; return false; }
            object range = null;
            try
            {
                if (entry.StoryType != 1) { message = "非正文目标没有可靠快照撤销。"; return false; }
                range = WordCom.Call(document, "Range", entry.Start, entry.End);
                if (!string.Equals(WordCom.ReadText(range), entry.After, StringComparison.Ordinal))
                { message = "目标文本已移动或再次修改，快照不能安全撤销。"; return false; }
                WordCom.Set(range, "Text", entry.Before);
                var readback = WordCom.ReadText(range);
                if (!string.Equals(readback, entry.Before, StringComparison.Ordinal))
                { message = "撤销读回结果不一致。"; return false; }
                entry.Consumed = true; message = "已根据文档快照撤销实际改动。"; return true;
            }
            catch (WordToolException ex) { message = ex.Message; return false; }
            catch (Exception ex) { message = ex.Message; return false; }
            finally { WordCom.Release(range); }
        }

        private static IntPtr Identity(object document)
        {
            if (!Marshal.IsComObject(document)) { return IntPtr.Zero; }
            var pointer = Marshal.GetIUnknownForObject(document);
            try { return pointer; }
            finally { Marshal.Release(pointer); }
        }
    }
}
