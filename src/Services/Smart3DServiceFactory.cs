using System;
using System.Runtime.InteropServices;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.ServiceManager;
using Ingr.SP3D.Content.ServiceManager;

namespace Smart3D.ExcelImport.Services
{
    /// <summary>
    /// Factory for accessing Smart3D V14 application services.
    /// Uses ServiceManager.GetService pattern for service resolution.
    /// </summary>
    public static class Smart3DServiceFactory
    {
        private static Application _cachedApplication;

        /// <summary>
        /// Gets the active Smart3D application instance via ServiceManager.
        /// </summary>
        public static Application GetApplication()
        {
            if (_cachedApplication != null)
                return _cachedApplication;

            try
            {
                // Smart3D V14: Get application via ServiceManager
                _cachedApplication = ServiceManager.GetService(ServiceType.Application) as Application;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Unable to connect to Smart3D V14 application. " +
                    "Ensure Smart3D is installed and running.", ex);
            }

            if (_cachedApplication == null)
                throw new InvalidOperationException("Failed to obtain Smart3D application reference.");

            return _cachedApplication;
        }

        /// <summary>
        /// Gets a specific service from the Smart3D ServiceManager.
        /// </summary>
        public static T GetService<T>() where T : class
        {
            try
            {
                // Try ServiceManager first (core services)
                var service = ServiceManager.GetService(typeof(T));
                if (service != null)
                    return service as T;

                // Try ContentServiceManager for content services
                var contentService = ContentServiceManager.GetService(typeof(T));
                return contentService as T;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Clears the cached application reference (useful for reconnection).
        /// </summary>
        public static void Reset()
        {
            if (_cachedApplication != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_cachedApplication);
                }
                catch { }
                _cachedApplication = null;
            }
        }

        /// <summary>
        /// Checks if Smart3D is currently accessible.
        /// </summary>
        public static bool IsAvailable()
        {
            try
            {
                var app = GetApplication();
                return app != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
