using System;
using System.Runtime.InteropServices;
using BenchmarkDotNet.Attributes;
using BenchmarksPerformanceIssues.ObjectPools.OfficeInterop;
using Microsoft.Office.Interop.Access;

namespace BenchmarksPerformanceIssues.Benchmarks.InteropInstanceCreation
{
    [RyuJitX64Job]
    [MeanColumn(), MedianColumn(), MaxColumn]
    [MemoryDiagnoser]
    [MinIterationCount(5), MaxIterationCount(10)]
    public class InteropInstanceCreationBenchmark
    {
        private readonly MicrosoftAccessObjectPool _pool;

        public InteropInstanceCreationBenchmark()
        {
            _pool = new MicrosoftAccessObjectPool(new MicrosoftAccessPooledObjectPolicy(quitQuitBeforeRelease: true), 2);
        }

        [Benchmark()]
        public void CreateInstanceUsingPool()
        {
            var instance = _pool.Get();

            _pool.Return(instance);
        }

        [Benchmark]
        public void CreateInstanceUsingActivator()
        {
            var instance =
                (Application)Activator.CreateInstance(
                    Marshal.GetTypeFromCLSID(new Guid("73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9")));

            instance.Quit(AcQuitOption.acQuitSaveAll);

            Marshal.ReleaseComObject(instance);
        }
        [Benchmark]
        public void CreateInstanceWithLateBinding()
        {
            Type t = Type.GetTypeFromProgID("Access.Application");
            var instance = (Application)Activator.CreateInstance(t);
            instance.Quit(AcQuitOption.acQuitSaveAll);

            Marshal.ReleaseComObject(instance);
        }

        [GlobalCleanup]
        public void CleanUp()
        {
            _pool.Dispose();
        }
    }
}
