using System;
using System.Runtime.InteropServices;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.ServiceManager;
using Ingr.SP3D.Content;
using Ingr.SP3D.UI;
using Smart3D.ExcelImport.Core;

namespace Smart3D.ExcelImport.Commands
{
    /// <summary>
    /// Base class for Smart3D V14 commands.
    /// Provides common infrastructure for command execution including
    /// application context access, transaction management, and error handling.
    /// </summary>
    public abstract class CommandBase : IDisposable
    {
        protected Application Application { get; private set; }
        protected IModel ActiveModel => Application?.ActiveProject?.ActiveModel;
        protected IObjectFactory ObjectFactory => ActiveModel?.ObjectFactory;

        private bool _disposed;

        /// <summary>
        /// Initializes the command with the current Smart3D application context.
        /// </summary>
        protected CommandBase()
        {
            try
            {
                // Smart3D application retrieved via ServiceManager
                Application = Smart3DApplicationHelper.ActiveApplication;
                if (Application == null)
                {
                    throw new InvalidOperationException(
                        "Failed to connect to Smart3D application. Ensure Smart3D is running.");
                }
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    "Failed to connect to Smart3D application. Ensure Smart3D is running.", ex);
            }
        }

        /// <summary>
        /// Execute the command logic. Override in derived classes.
        /// </summary>
        public abstract void Execute();

        /// <summary>
        /// Gets a specific service from the Smart3D ServiceManager.
        /// </summary>
        protected T GetService<T>() where T : class
        {
            return ServiceManager.GetService(typeof(T)) as T;
        }

        #region IDisposable
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Release managed resources
                }
                _disposed = true;
            }
        }
        #endregion
    }
}
