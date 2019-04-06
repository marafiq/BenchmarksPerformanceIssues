using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Running;

namespace BenchmarksPerformanceIssues
{
    class Program
    {
        static void Main(string[] args)
        {
            BenchmarkSwitcher
                .FromAssembly(typeof(Program).Assembly)
                .Run(args, DefaultConfig.Instance);
        }
    }
}
