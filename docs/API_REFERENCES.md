# Smart3D V14 API References

## Required DLLs

| # | DLL | Namespace | Purpose |
|---|-----|-----------|---------|
| 1 | Ingr.SP3D.Common.Middle | `Ingr.SP3D.Common.Middle` | Core types, vectors, matrices, service access |
| 2 | Ingr.SP3D.Common.Middle.ServiceManager | `Ingr.SP3D.Common.Middle.ServiceManager` | Application services (filtering, transactions) |
| 3 | Ingr.SP3D.Content | `Ingr.SP3D.Content` | Model objects, relations, property manipulation |
| 4 | Ingr.SP3D.Content.DataAccess | `Ingr.SP3D.Content.DataAccess` | Database queries, object retrieval |
| 5 | Ingr.SP3D.Content.ServiceManager | `Ingr.SP3D.Content.ServiceManager` | Content-specific services |
| 6 | Ingr.SP3D.UI | `Ingr.SP3D.UI` | Command framework, user interface |
| 7 | Ingr.SP3D.SystemsAndSpecifications | `Ingr.SP3D.SystemsAndSpecifications` | Spec-driven design, spec objects |

## Key Interfaces

### Application Root
```csharp
// Get the active application
Application app = ServiceManager.GetService(ServiceType.Application) as Application;
IModel model = app.ActiveProject.ActiveModel;
```

### Model Access
```csharp
IModel model = ...;                    // Active model
ISystemRoot systems = model.Systems;   // System hierarchy
IObject obj = ...;                     // Any model object
IPropertyValues props = obj.PropertyValues;  // Property bag
```

### Filtering Service
```csharp
IFiltering filter = ServiceManager.GetService(ServiceType.Filtering) as IFiltering;
IObjectCollection results = filter.GetObjects(model, "Name = '100A-P-101'");
```

### Property Manipulation
```csharp
var propValues = obj.PropertyValues;

// Read
string desc = propValues["Description"].ToString();

// Write
propValues["Description"] = "New Value";
propValues["DesignTemp"] = 150.0;        // Numeric
propValues["IsInsulated"] = true;        // Boolean

// Commit
obj.Update();
```

### Hierarchy Traversal
```csharp
// Site > Zone > System > Component
foreach (ISystem system in model.Systems)
{
    foreach (IObject member in system.Members)
    {
        // Process member
    }
    foreach (ISystem child in system.ChildSystems)
    {
        // Recurse
    }
}
```

### Transaction Management
```csharp
ITransaction txn = ServiceManager.GetService(ServiceType.Transaction) as ITransaction;
txn.StartTransaction();
try
{
    // Perform updates
    txn.CommitTransaction();
}
catch
{
    txn.RollbackTransaction();
}
```

### COM Interop Best Practices
```csharp
// Always release COM objects
Marshal.ReleaseComObject(comObj);

// Use try-finally for safety
IObject obj = null;
try
{
    obj = GetObject();
    // Use obj
}
finally
{
    if (obj != null) Marshal.ReleaseComObject(obj);
}
```

## Object Type Hierarchy

```
IObject (base)
├── ISystem (containers: Site, Zone, System)
│   └── Members: IObjectCollection
├── IPipingObject (piping components)
│   ├── IPipeRun
│   ├── IPipeLine
│   ├── IPipeFitting (Valve, Elbow, Tee, Reducer)
│   └── IPipeNozzle
├── IEquipment
│   └── IEquipmentComponent
├── IStructuralObject
│   └── IStructMember
├── IInstrument
└── IHangerSupport
```

## Property Value Types

| Property Pattern | .NET Type | Example |
|-----------------|-----------|---------|
| Name, Description, Type | `string` | `"Process Line"` |
| Count, Quantity | `int` | `42` |
| Temp, Pressure, Weight, Size | `double` | `150.5` |
| Is*, Has*, Enabled | `bool` | `true` |
| *Date, *Time | `DateTime` | `2026-06-27` |

## Catalog/Spec Access
```csharp
// Access catalog for spec information
var catalog = model.Catalog;
var spec = catalog.GetSpec("A1A");  // Get pipe spec
var components = spec.Components;    // Components in spec
```

## Excel Import — Complete Property Update Flow

```
1. Get Filtering service
2. Build filter: "Name = 'ObjectName' AND ObjectType = 'Type'"
3. Execute filter → get IObject
4. Access: obj.PropertyValues
5. Coerce value to target type
6: Set: propValues["AttributeName"] = value
7. Call: obj.Update()
8. Release: Marshal.ReleaseComObject(obj)
```
