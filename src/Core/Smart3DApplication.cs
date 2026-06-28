// ============================================================================
// Smart3D V14 - Excel Property Import
// File: Core/Smart3DApplication.cs
// Description: Helper class to access Smart3D application root and services
// ============================================================================

using System;
using System.Runtime.InteropServices;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.ServiceManager;
using Ingr.SP3D.Content.ServiceManager;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Provides access to the Smart3D application root and key services
    /// </summary>
    public static class Smart3DApplicationHelper
    {
        /// <summary>
        /// Gets the active Smart3D application instance
        /// </summary>
        public static Application ActiveApplication
        {
            get
            {
                try
                {
                    return ServiceManager.GetService(ServiceType.Application) as Application;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the active model in Smart3D
        /// </summary>
        public static IModel ActiveModel
        {
            get
            {
                try
                {
                    var app = ActiveApplication;
                    return app?.ActiveProject?.ActiveModel;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the Filtering service for object lookup
        /// </summary>
        public static IFiltering FilteringService
        {
            get
            {
                try
                {
                    return ServiceManager.GetService(ServiceType.Filtering) as IFiltering;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the Transaction service for batch operations
        /// </summary>
        public static ITransaction TransactionService
        {
            get
            {
                try
                {
                    return ServiceManager.GetService(ServiceType.Transaction) as ITransaction;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the Content DataAccess service
        /// </summary>
        public static IDataAccess DataAccessService
        {
            get
            {
                try
                {
                    return ContentServiceManager.GetService(ServiceType.DataAccess) as IDataAccess;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Safely releases a COM object
        /// </summary>
        public static void ReleaseComObject(object comObject)
        {
            if (comObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch (Exception)
                {
                    // Best-effort release
                }
            }
        }

        /// <summary>
        /// Executes an action within a Smart3D transaction
        /// </summary>
        public static T ExecuteInTransaction<T>(Func<T> action, ILogger logger)
        {
            var txn = TransactionService;
            if (txn == null)
            {
                logger?.Error("Transaction service not available");
                return default;
            }

            try
            {
                txn.StartTransaction();
                T result = action();
                txn.CommitTransaction();
                return result;
            }
            catch (Exception ex)
            {
                logger?.Error(ex, "Transaction failed, rolling back");
                try { txn.RollbackTransaction(); } catch { }
                throw;
            }
        }
    }
}
