using System;
using System.Runtime.InteropServices;
using Microsoft.Extensions.ObjectPool;
using Microsoft.Office.Interop.Access;

namespace BenchmarksPerformanceIssues.ObjectPools.OfficeInterop
{
    public class MicrosoftAccessPooledObjectPolicy : PooledObjectPolicy<Application>
    {
        private readonly Guid _microsoftAccessClassGuid = new Guid("73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9");
        private readonly bool _quitBeforeRelease;
        public MicrosoftAccessPooledObjectPolicy(bool quitQuitBeforeRelease = true)
        {
            _quitBeforeRelease = quitQuitBeforeRelease;
        }
        #region Overrides of PooledObjectPolicy<Application>

        /// <summary>
        /// Responsible for creating the Application object
        /// </summary>
        /// <returns></returns>
        public override Application Create()
        {
            return (Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(_microsoftAccessClassGuid));
        }

        /// <summary>
        /// Return the object to pool, actually disposing the com object.
        /// Quit of com object can be configured using constructor, otherwise call quit before calling return.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Return(Application obj)
        {
            if (_quitBeforeRelease) obj.Quit(AcQuitOption.acQuitSaveAll);
            Marshal.ReleaseComObject(obj);
            return false;
        }

        #endregion
    }
}
