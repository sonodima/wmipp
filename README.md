<h1 align="center">WMI++ 🤕</h1>

<div align="center">
  <img src="https://img.shields.io/badge/c%2B%2B-20-orange"/>
  <img src="https://img.shields.io/badge/license-MIT-blue.svg"/>
</div>

<br>

> WMI++ is a tiny-but-mighty header-only C++20 library that takes away the pain of dealing with the Windows Management Instrumentation (WMI).

> If you've ever had the necessity to use WMI, you know it can be like wrestling with a grumpy elephant.
> 
> WMI++ swoops in like a nimble superhero to save the day, providing a clean and intuitive interface for your WMI needs.

## Usage

#### Minimal Example

Retrieve a specific property from a WMI query using WMI++.

The code executes a query that selects the `Name` property from the `Win32_Processor` class.

The result is accessed and stored as a `std::optional<std::string>` in the `value` variable.

```cpp
#include <wmipp/wmipp.hxx>

const auto value = wmipp::Interface::Create()
  ->ExecuteQuery(L"SELECT Name FROM Win32_Processor")
  .GetProperty<std::string>(L"Name");
```

#### Custom Path

Connect to a specific WMI namespace by providing a custom path to the `wmipp::Interface::Create` method.

Replace `"CUSTOM_PATH_HERE"` with the desired custom path, such as the namespace or machine path.

```cpp
#include <wmipp/wmipp.hxx>

const auto iface = wmipp::Interface::Create("CUSTOM_PATH_HERE");
```

[MSDN](https://learn.microsoft.com/en-us/windows/win32/wmisdk/describing-the-location-of-a-wmi-object)

#### Object Iteration

Iterate over multiple WMI objects returned by a query.

The code executes a query that selects the Model property from the `Win32_DiskDrive`.

Because this query may return more than one result _(with disks > 1)_, you can choose to iterate
all the returned objects in the result with range-based loops.

Within the loop, retrieve the `Model` property value of each object using the `GetProperty` function and store it in the `model` variable.

```cpp
#include <wmipp/wmipp.hxx>

for (const auto& obj : wmipp::Interface::Create()->ExecuteQuery(L"SELECT Model FROM Win32_DiskDrive")) {
  const auto model = obj.GetProperty<std::string>(L"Model");
}
```

#### Reutilizing Interfaces

If you need to perform more than one query on the same path, you can utilize the same `Interface` to avoid creating
a new connection for each query.

The `Interface` instance will automatically get uninitialized when all other `Objects` and `QueryResults`
using it go out of scope.

```cpp
#include <wmipp/wmipp.hxx>

const auto iface = wmipp::Interface::Create();
const auto cpu_name = iface->ExecuteQuery(L"SELECT Name FROM Win32_Processor")
  .GetProperty<std::string>(L"Name");
const auto user_name = iface->ExecuteQuery(L"SELECT UserName FROM Win32_ComputerSystem")
  .GetProperty<std::string>(L"UserName");
```

## About Typing

Currently there is support for the majority of the types you would usually need to query.
This is done by building a __variant_t__ and delegating it to handle the conversions.
Furthermore, strings are converted from bstr_ts, and arrays are built from CComSafeArrays.

However, support is currently not guaranteed for all types that can be present in VARIANTs.
