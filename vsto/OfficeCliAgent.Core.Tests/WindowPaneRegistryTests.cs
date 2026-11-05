using System;
using OfficeCliAgent.Core;
using Xunit;

namespace OfficeCliAgent.Core.Tests
{
    public class WindowPaneRegistryTests
    {
        private sealed class Owner : IDisposable
        {
            public bool Disposed;
            public void Dispose() => Disposed = true;
        }

        private sealed class Pane
        {
            public bool Closed;
        }

        [Fact]
        public void FindsThePaneOfAWindowByHandle()
        {
            var registry = new WindowPaneRegistry<Pane>();
            var pane = new Pane();
            registry.Add(0x10, pane, new Owner());

            Assert.True(registry.TryGet(0x10, out var found));
            Assert.Same(pane, found);
            Assert.False(registry.TryGet(0x20, out _));
        }

        [Fact]
        public void PruneDropsPanesWhoseWindowClosedAndDisposesTheirOwner()
        {
            var registry = new WindowPaneRegistry<Pane>();
            var closedOwner = new Owner();
            var liveOwner = new Owner();
            registry.Add(1, new Pane { Closed = true }, closedOwner);
            registry.Add(2, new Pane(), liveOwner);

            var dropped = registry.Prune(p => !p.Closed);

            Assert.Single(dropped);
            Assert.True(dropped[0].Closed);

            Assert.True(closedOwner.Disposed);
            Assert.False(liveOwner.Disposed);
            Assert.False(registry.TryGet(1, out _));
            Assert.True(registry.TryGet(2, out _));
        }

        [Fact]
        public void PruneTreatsADisposedPaneAsDead()
        {
            var registry = new WindowPaneRegistry<Pane>();
            var owner = new Owner();
            registry.Add(1, new Pane(), owner);

            Assert.Single(registry.Prune(_ => throw new ObjectDisposedException("pane")));
            Assert.True(owner.Disposed);
            Assert.Equal(0, registry.Count);
        }

        [Fact]
        public void ReplacingAWindowsPaneDisposesThePreviousOwner()
        {
            var registry = new WindowPaneRegistry<Pane>();
            var first = new Owner();
            registry.Add(1, new Pane(), first);
            registry.Add(1, new Pane(), new Owner());

            Assert.True(first.Disposed);
            Assert.Equal(1, registry.Count);
        }

        [Fact]
        public void DisposeDisposesEveryOwner()
        {
            var registry = new WindowPaneRegistry<Pane>();
            var a = new Owner();
            var b = new Owner();
            registry.Add(1, new Pane(), a);
            registry.Add(2, new Pane(), b);

            registry.Dispose();

            Assert.True(a.Disposed);
            Assert.True(b.Disposed);
            Assert.Equal(0, registry.Count);
        }
    }
}
