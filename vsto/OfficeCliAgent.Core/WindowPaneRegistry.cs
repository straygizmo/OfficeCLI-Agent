using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeCliAgent.Core
{
    /// <summary>
    /// The chat panes an add-in has created, one per document window, keyed by the window's HWND.
    ///
    /// Office 2013+ opens every document in its own top-level window (SDI), and a VSTO custom
    /// task pane belongs to exactly one window. A pane created once at startup therefore dies
    /// together with the blank "Document1" that Office closes when an existing file is opened
    /// from it, and the ribbon toggle has nothing left to show. Each add-in instead asks this
    /// registry for the active window's pane and creates one when it is missing.
    /// </summary>
    /// <typeparam name="TPane">The VSTO CustomTaskPane type (kept generic so Core stays VSTO-free).</typeparam>
    public sealed class WindowPaneRegistry<TPane> : IDisposable where TPane : class
    {
        private sealed class Entry
        {
            public TPane Pane;
            public IDisposable Owner;
        }

        private readonly Dictionary<int, Entry> _entries = new Dictionary<int, Entry>();

        public int Count => _entries.Count;

        public bool TryGet(int hwnd, out TPane pane)
        {
            if (_entries.TryGetValue(hwnd, out var entry))
            {
                pane = entry.Pane;
                return true;
            }
            pane = null;
            return false;
        }

        /// <param name="owner">Disposed together with the pane's registry entry (the controller behind the pane).</param>
        public void Add(int hwnd, TPane pane, IDisposable owner)
        {
            if (pane == null) throw new ArgumentNullException(nameof(pane));
            Remove(hwnd);
            _entries[hwnd] = new Entry { Pane = pane, Owner = owner };
        }

        /// <summary>Forget the pane of <paramref name="hwnd"/> and dispose its owner. Returns false when unknown.</summary>
        public bool Remove(int hwnd)
        {
            if (!_entries.TryGetValue(hwnd, out var entry)) return false;
            _entries.Remove(hwnd);
            DisposeOwner(entry);
            return true;
        }

        /// <summary>
        /// Drop every pane that <paramref name="isAlive"/> rejects, disposing its owner, and return
        /// the dropped panes so the add-in can remove them from CustomTaskPanes as well.
        /// </summary>
        public IReadOnlyList<TPane> Prune(Func<TPane, bool> isAlive)
        {
            var dead = _entries.Where(kv => !IsAlive(isAlive, kv.Value.Pane)).ToList();
            foreach (var kv in dead)
            {
                Log.Write($"panes: window {kv.Key} is gone, dropping its pane");
                Remove(kv.Key);
            }
            return dead.Select(kv => kv.Value.Pane).ToList();
        }

        private static bool IsAlive(Func<TPane, bool> isAlive, TPane pane)
        {
            try { return isAlive(pane); }
            catch (ObjectDisposedException) { return false; }
        }

        private static void DisposeOwner(Entry entry)
        {
            try { entry.Owner?.Dispose(); }
            catch (Exception ex) { Log.Write("panes: dispose failed: " + ex.Message); }
        }

        public void Dispose()
        {
            foreach (var entry in _entries.Values.ToList()) DisposeOwner(entry);
            _entries.Clear();
        }
    }
}
