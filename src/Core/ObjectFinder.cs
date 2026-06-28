using System;
using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.ServiceManager;
using Ingr.SP3D.Content;
using Ingr.SP3D.Content.DataAccess;
using Smart3D.ExcelImport.Services;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Finds objects in the Smart3D model hierarchy using the Filtering service.
    /// Supports searching by name, type, and system hierarchy.
    /// </summary>
    public sealed class ObjectFinder
    {
        private readonly Application _application;
        private readonly LoggingService _logger;
        private readonly IFiltering _filteringService;

        public ObjectFinder(Application application, LoggingService logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            _filteringService = Smart3DApplicationHelper.FilteringService;

            if (_filteringService == null)
            {
                // Fallback: use DataAccess service for object queries
                _logger.LogWarning("Filtering service not available, will use DataAccess fallback.");
            }
        }

        /// <summary>
        /// Finds a single object by its name in the model.
        /// </summary>
        public IObject FindObjectByName(string objectName, string objectType = null)
        {
            if (string.IsNullOrWhiteSpace(objectName))
                return null;

            _logger.LogInfo($"Searching for object: '{objectName}' (Type: {objectType ?? "Any"})");

            try
            {
                // Strategy 1: Use Filtering service with name filter
                if (_filteringService != null)
                {
                    var model = _application.ActiveProject?.ActiveModel;
                    if (model != null)
                    {
                        string filterCriteria = $"Name = '{objectName.Replace("'", "''")}'";
                        if (!string.IsNullOrEmpty(objectType))
                        {
                            filterCriteria += $" AND Type = '{objectType}'";
                        }

                        var objects = _filteringService.GetObjects(model, filterCriteria);
                        if (objects != null && objects.Count > 0)
                        {
                            _logger.LogInfo($"Found object '{objectName}' via Filtering service.");
                            return objects[0];
                        }
                    }
                }

                // Strategy 2: Use DataAccess with SQL-like query
                var obj = FindViaDataAccess(objectName, objectType);
                if (obj != null)
                {
                    _logger.LogInfo($"Found object '{objectName}' via DataAccess.");
                    return obj;
                }

                // Strategy 3: Search through system hierarchy
                obj = FindInHierarchy(objectName, objectType);
                if (obj != null)
                {
                    _logger.LogInfo($"Found object '{objectName}' via hierarchy traversal.");
                    return obj;
                }

                _logger.LogWarning($"Object '{objectName}' not found in model.");
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error finding object '{objectName}'", ex);
                return null;
            }
        }

        /// <summary>
        /// Finds multiple objects matching the given names.
        /// </summary>
        public Dictionary<string, IObject> FindMultipleObjects(
            IEnumerable<string> objectNames, string objectType = null)
        {
            var results = new Dictionary<string, IObject>(StringComparer.OrdinalIgnoreCase);
            var nameList = objectNames.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            _logger.LogInfo($"Bulk searching for {nameList.Count} objects...");

            try
            {
                // Use filtering service for bulk search when possible
                if (_filteringService != null)
                {
                    var model = _application.ActiveProject?.ActiveModel;
                    if (model != null)
                    {
                        // Build OR-based filter for multiple names
                        var nameConditions = string.Join(" OR ", nameList.Select(n => $"Name = '{n.Replace("'", "''")}'"));
                        string filterCriteria = $"({nameConditions})";
                        if (!string.IsNullOrEmpty(objectType))
                        {
                            filterCriteria += $" AND Type = '{objectType}'";
                        }

                        var objects = _filteringService.GetObjects(model, filterCriteria);
                        if (objects != null)
                        {
                            foreach (var obj in objects)
                            {
                                if (obj != null && !string.IsNullOrEmpty(obj.Name) && !results.ContainsKey(obj.Name))
                                    results[obj.Name] = obj;
                            }
                        }
                    }
                }

                // Find missing objects individually
                var missingNames = nameList.Where(n => !results.ContainsKey(n)).ToList();
                foreach (var name in missingNames)
                {
                    var found = FindObjectByName(name, objectType);
                    if (found != null)
                        results[name] = found;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Error during bulk object search", ex);
            }

            _logger.LogInfo($"Found {results.Count} of {nameList.Count} requested objects.");
            return results;
        }

        /// <summary>
        /// Finds objects using DataAccess service as a fallback.
        /// </summary>
        private IObject FindViaDataAccess(string objectName, string objectType)
        {
            try
            {
                var dataAccess = ServiceManager.GetService(ServiceType.DataAccess) as IDataAccess;
                if (dataAccess == null) return null;

                // Query the object table by name
                var query = $"SELECT * FROM Object WHERE Name = '{objectName.Replace("'", "''")}'";
                if (!string.IsNullOrEmpty(objectType))
                {
                    query += $" AND Type = '{objectType}'";
                }

                var resultSet = dataAccess.ExecuteQuery(query);
                if (resultSet?.Tables.Count > 0 && resultSet.Tables[0].Rows.Count > 0)
                {
                    var objectId = Convert.ToInt64(resultSet.Tables[0].Rows[0]["ObjectId"]);
                    return _application.ActiveProject?.ActiveModel?.GetObjectById(objectId);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"DataAccess fallback failed for '{objectName}': {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Searches through the system hierarchy for an object by name.
        /// </summary>
        private IObject FindInHierarchy(string objectName, string objectType)
        {
            try
            {
                var model = _application.ActiveProject?.ActiveModel;
                if (model == null) return null;

                // Traverse the system hierarchy using Systems property
                var rootSystems = model.Systems;
                if (rootSystems == null) return null;

                foreach (var system in rootSystems)
                {
                    var found = SearchSystemRecursive(system, objectName, objectType);
                    if (found != null) return found;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Hierarchy search failed for '{objectName}': {ex.Message}");
            }

            return null;
        }

        private IObject SearchSystemRecursive(ISystem system, string objectName, string objectType)
        {
            if (system == null) return null;

            // Check members of this system
            var members = system.Members;
            if (members != null)
            {
                foreach (var member in members)
                {
                    if (member is IObject obj &&
                        obj.Name.Equals(objectName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(objectType) ||
                            obj.Type.ToString().Equals(objectType, StringComparison.OrdinalIgnoreCase))
                        {
                            return obj;
                        }
                    }
                }
            }

            // Recurse into child systems
            var childSystems = system.ChildSystems;
            if (childSystems != null)
            {
                foreach (var child in childSystems)
                {
                    if (child is ISystem subSystem)
                    {
                        var found = SearchSystemRecursive(subSystem, objectName, objectType);
                        if (found != null) return found;
                    }
                }
            }

            return null;
        }
    }
}
