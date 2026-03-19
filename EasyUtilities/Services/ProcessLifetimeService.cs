using System.Collections.Concurrent;
using System.Diagnostics;

namespace EasyUtilities.Services;

public static class ProcessLifetimeService
{
    private static readonly ConcurrentDictionary<int, Process> TrackedProcesses = new();
    private static bool _suppressSiblingShutdown;

    public static void Register(Process? process)
    {
        if (process is null)
        {
            return;
        }

        if (!TrackedProcesses.TryAdd(process.Id, process))
        {
            return;
        }

        try
        {
            process.EnableRaisingEvents = true;
            process.Exited += (_, _) => Remove(process.Id);
        }
        catch
        {
            Remove(process.Id);
        }
    }

    public static void SuppressSiblingShutdownForCurrentExit()
    {
        _suppressSiblingShutdown = true;
    }

    public static void CleanupOnExit()
    {
        KillTrackedProcesses();

        if (_suppressSiblingShutdown)
        {
            return;
        }

        KillSiblingInstances();
    }

    private static void KillTrackedProcesses()
    {
        foreach (var pair in TrackedProcesses.ToArray())
        {
            var process = pair.Value;

            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                    process.WaitForExit(3000);
                }
            }
            catch
            {
                // no-op
            }
            finally
            {
                Remove(pair.Key);
            }
        }
    }

    private static void KillSiblingInstances()
    {
        var current = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(current))
        {
            return;
        }

        var processName = Path.GetFileNameWithoutExtension(current);
        foreach (var process in Process.GetProcessesByName(processName))
        {
            try
            {
                if (process.Id == Environment.ProcessId)
                {
                    continue;
                }

                var sameExecutable = false;
                try
                {
                    sameExecutable = string.Equals(process.MainModule?.FileName, current, StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    // Access denied for elevated/system process - ignore.
                }

                if (!sameExecutable || process.HasExited)
                {
                    continue;
                }

                process.Kill(entireProcessTree: true);
                process.WaitForExit(3000);
            }
            catch
            {
                // no-op
            }
            finally
            {
                process.Dispose();
            }
        }
    }

    private static void Remove(int pid)
    {
        if (TrackedProcesses.TryRemove(pid, out var process))
        {
            process.Dispose();
        }
    }
}
