using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.ServiceManager;
using Ingr.SP3D.Content;
using Smart3D.ExcelImport.Services;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Manages Smart3D transactions for batch operations.
    /// Ensures data consistency and provides rollback capability.
    /// </summary>
    public sealed class TransactionManager
    {
        private readonly Application _application;
        private readonly LoggingService _logger;

        public TransactionManager(Application application, LoggingService logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Begins a new transaction with the specified name using TransactionService.
        /// </summary>
        public ITransaction BeginTransaction(string transactionName)
        {
            _logger.LogInfo($"Starting transaction: {transactionName}");

            var transaction = ServiceManager.GetService(ServiceType.Transaction) as ITransaction;
            if (transaction == null)
                throw new InvalidOperationException("Transaction service not available.");

            transaction.StartTransaction();
            _logger.LogInfo($"Transaction '{transactionName}' started successfully.");

            return transaction;
        }

        /// <summary>
        /// Executes an action within a transaction scope with automatic rollback on failure.
        /// </summary>
        public T ExecuteInTransaction<T>(string name, Func<ITransaction, T> action)
        {
            var transaction = BeginTransaction(name);
            try
            {
                var result = action(transaction);
                transaction.CommitTransaction();
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Transaction '{name}' failed, rolling back.", ex);
                try { transaction.RollbackTransaction(); } catch { }
                throw;
            }
        }

        /// <summary>
        /// Executes an action within a transaction scope (no return value).
        /// </summary>
        public void ExecuteInTransaction(string name, Action<ITransaction> action)
        {
            var transaction = BeginTransaction(name);
            try
            {
                action(transaction);
                transaction.CommitTransaction();
            }
            catch (Exception ex)
            {
                _logger.LogError($"Transaction '{name}' failed, rolling back.", ex);
                try { transaction.RollbackTransaction(); } catch { }
                throw;
            }
        }
    }
}
