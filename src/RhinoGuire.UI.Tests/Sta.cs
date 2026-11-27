using System.Windows;
using System.Windows.Threading;

namespace RhinoGuire.UI.Tests;

/// <summary>
/// Runs test code on one long-lived STA thread that owns a WPF <see cref="Application"/>, so pack URIs
/// (embedded dictionaries and fonts) resolve as they do inside Rhino.
/// </summary>
internal static class Sta
{
    private static readonly Lazy<Dispatcher> Dispatcher = new(Start, LazyThreadSafetyMode.ExecutionAndPublication);

    private static Dispatcher Start()
    {
        Dispatcher? dispatcher = null;
        using var ready = new ManualResetEventSlim();
        var thread = new Thread(() =>
        {
            if (Application.Current == null)
                _ = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
            dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher;
            ready.Set();
            System.Windows.Threading.Dispatcher.Run();
        })
        {
            IsBackground = true,
            Name = "RhinoGuire.UI.Tests STA",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        ready.Wait();
        return dispatcher!;
    }

    public static void Run(Action action) => Dispatcher.Value.Invoke(action);

    public static T Run<T>(Func<T> func) => Dispatcher.Value.Invoke(func);
}

/// <summary>Repository paths for tests that read source files.</summary>
internal static class Repo
{
    public static string Root { get; } = FindRoot();

    private static string FindRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "src", "RhinoGuire.sln")))
            dir = dir.Parent;
        return dir?.FullName ?? throw new InvalidOperationException("Repository root (src/RhinoGuire.sln) not found.");
    }

    public static string Path_(params string[] parts) => Path.Combine(new[] { Root }.Concat(parts).ToArray());
}
