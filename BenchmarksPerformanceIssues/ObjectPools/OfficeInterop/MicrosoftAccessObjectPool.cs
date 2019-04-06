using System;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using Microsoft.Extensions.ObjectPool;
using Microsoft.Office.Interop.Access;

namespace BenchmarksPerformanceIssues.ObjectPools.OfficeInterop
{
    public class MicrosoftAccessObjectPool : ObjectPool<Application>,IDisposable
    {
        private readonly IPooledObjectPolicy<Application> _policy;
        readonly ConcurrentQueue<Application> _concurrentQueue = new ConcurrentQueue<Application>();
        private readonly int _maxInstancesInPool;

        public MicrosoftAccessObjectPool(IPooledObjectPolicy<Application> policy) : this(policy, Environment.ProcessorCount)
        {

        }

        public MicrosoftAccessObjectPool(IPooledObjectPolicy<Application> policy, int maxInstancesInPool)
        {
            _policy = policy;
            _maxInstancesInPool = maxInstancesInPool;

            FillQueue();
        }


        #region Overrides of ObjectPool<Application>
        /// <summary>
        /// Get object from pool, if not available create as per supplied policy.
        /// It uses concurrent queue as backing store, as soon queue is empty, queue will be filled again.
        /// </summary>
        /// <returns></returns>
        public override Application Get()
        {
            var x = _concurrentQueue.TryDequeue(out Application application);
            Task.Run(() =>
            {
                if (_concurrentQueue.Count == 0)
                    FillQueue();
            }).ConfigureAwait(false);


            return x ? application : _policy.Create();
        }

        /// <summary>
        /// Return the object to pool, return of supplied policy will be used.
        /// Return is called Fire & Forget way, due to expensive nature of disposing com object.
        /// </summary>
        /// <param name="obj"></param>
        public override void Return(Application obj)
        {
            Task.Run(() => _policy.Return(obj)).ConfigureAwait(false);
        }

        #endregion
        

        private void FillQueue()
        {
            Parallel.For(0, _maxInstancesInPool, new ParallelOptions() { MaxDegreeOfParallelism = 2 }, i =>
                {
                    _concurrentQueue.Enqueue(_policy.Create());
                });
        }

        #region Implementation of IDisposable

        /// <summary>
        /// Dispose all objects, in backing store, so process of com object do not hang in memory.
        /// </summary>
        public void Dispose()
        {
            for (int i = 0; i < _concurrentQueue.Count; i++)
            {
                var x = _concurrentQueue.TryDequeue(out Application application);
                if (x == true) _policy.Return(application);
            }
        }

        #endregion
    }
}
