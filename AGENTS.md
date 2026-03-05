# AGENTS.md

# 0. Requirements

Always show your **Plan** before making any changes.

# 1. Project Overview

This project is an internal desktop utility platform designed for use in a **restricted corporate environment** where:

* Internet access is unavailable
* Software installation privileges are limited
* Enterprise security policies are enforced

The application provides a **modular desktop tool framework** where functionality can be extended through **plugin modules**.

The system prioritizes:

* portability
* reliability
* security compliance
* ease of extension

The tool is primarily intended for **internal productivity automation and utilities**.

---

# 2. Operating Environment

## Target Operating System

* Windows 10 or Windows 11 (enterprise managed)

## Network Conditions

* No internet access
* No external service connectivity

## Installation Restrictions

* Installation privileges are limited
* Software must run without installers

Applications must run directly from a copied directory.

---

# 3. Runtime Environment

Target machines include:

* **.NET Framework 4.8.1**

Confirmed via registry release value:

```
0x82348
```

Therefore the project must:

* Target **.NET Framework 4.8**
* Avoid modern .NET runtimes (.NET 6+)
* Avoid requiring runtime installation

---

# 4. Development Model

## Development Machine

Development occurs on a personal machine with full development tooling available.

Compiled binaries will be transferred manually to the corporate machine.

## Deployment Method

Deployment must be **portable**.

Applications must run by copying a directory containing all necessary files.

Example layout:

```
AppRoot/
 ├─ App.exe
 ├─ Contracts.dll
 ├─ Modules/
 │   ├─ ExampleModule.dll
 │   ├─ FileTools.dll
 │   └─ DataUtilities.dll
 ├─ Config/
 └─ Logs/
```

No installer or setup program should be required.

---

# 5. Security and Compliance Constraints

The application must comply with corporate security constraints.

The application must NOT:

* access the internet
* attempt automatic updates
* install additional runtimes
* require administrator privileges

The application should also avoid:

* Windows registry usage
* writing files outside the application directory
* system-wide configuration

All runtime data must remain within the application folder.

---

# 6. Code Signing

The developer possesses a valid code signing certificate.

All distributed binaries should be signed:

* Main executable
* Plugin DLLs

Signing helps prevent issues with enterprise endpoint security systems.

---

# 7. Technology Stack

Language:

* C#

Framework:

* .NET Framework 4.8

User Interface:

* WPF (Windows Presentation Foundation)

Reasons for choosing WPF:

* Native Windows UI
* Mature enterprise support
* No external dependencies
* Good separation between UI and logic

---

# 8. Architectural Overview

The system follows a **host + plugin architecture**.

Core components:

1. Host Application
2. Contracts Assembly
3. Plugin Modules

---

# 9. Repository Structure

The repository should follow a clear modular layout.

Recommended structure:

```
/src
  /HostApp
  /Contracts
  /Plugins
      /ExamplePlugin
      /FileUtilities
      /DataTools

/build
/docs

AGENTS.md
README.md
```

Explanation:

HostApp
Contains the WPF shell and plugin loader.

Contracts
Contains shared interfaces used by both host and plugins.

Plugins
Contains independent modules implementing features.

---

# 10. Host Application Responsibilities

The host application acts as the runtime container.

Responsibilities include:

* plugin discovery
* plugin loading
* plugin lifecycle management
* UI shell
* configuration management
* logging

The host must remain **stable and lightweight**.

---

# 11. Contracts Assembly

The Contracts assembly defines the shared interfaces used between the host and plugin modules.

This assembly must contain:

* plugin interfaces
* shared data models
* service abstractions

The Contracts project must NOT contain:

* UI logic
* business logic
* runtime behavior

It is purely an **API contract layer**.

---

# 12. Plugin Architecture

Plugins are implemented as separate assemblies.

Characteristics:

* each plugin is a DLL
* plugins reside inside `/Modules`
* plugins implement interfaces defined in Contracts

Plugins must be **independent modules**.

Adding a plugin should require only copying a DLL to the Modules directory.

---

# 13. Plugin Discovery Behavior

At application startup:

1. The host scans the `/Modules` directory.
2. All assemblies inside the directory are inspected.
3. Types implementing the plugin interface are discovered.
4. Valid plugins are instantiated and registered.

The loading system must be **fault tolerant**.

If a plugin fails to load:

* the error must be logged
* the host application must continue running

---

# 14. Plugin Interface Specification

All plugins must implement a common interface.

Minimum responsibilities for a plugin:

* expose plugin name
* expose plugin description
* provide an initialization method
* optionally provide UI integration

The plugin interface should remain **small and stable**.

Breaking changes must be avoided.

If future evolution is required, versioned interfaces should be introduced (for example `IPluginV2`).

---

# 15. Logging System

Logging must be simple and file-based.

Requirements:

* logs stored inside `/Logs`
* readable text format
* timestamped entries

Logs should capture:

* application startup
* plugin discovery
* plugin load success or failure
* runtime errors

Logging must never crash the application.

---

# 16. Configuration

Configuration files must be stored in `/Config`.

Preferred formats:

* JSON
* XML

Configuration must not depend on Windows registry.

Configuration files should be editable by advanced users.

---

# 17. Dependency Policy

Because deployment occurs in an offline environment:

* minimize external dependencies
* prefer built-in .NET libraries
* bundle any required third-party libraries

The project must not rely on internet package restoration during deployment.

---

# 18. UI Design Principles

The UI should prioritize:

* clarity
* simplicity
* responsiveness

The host UI should function as a **tool shell**.

Plugin modules may provide UI elements or commands that integrate into the host interface.

The UI should avoid unnecessary complexity.

---

# 19. Stability Requirements

The system must prioritize stability.

Requirements:

* host application must never crash due to plugin failure
* plugin loading must be isolated
* errors must be logged instead of terminating execution

The application should fail **gracefully** whenever possible.

---

# 20. Coding Conventions for Agents

When generating code, follow these rules:

Prefer:

* clear naming
* small focused classes
* minimal dependencies
* simple architecture

Avoid:

* unnecessary frameworks
* large dependency chains
* runtime reflection beyond plugin discovery
* network calls

Generated code should be:

* readable
* maintainable
* compatible with .NET Framework 4.8

---

# 21. Extensibility Goals

The architecture must allow new features to be added by:

1. Creating a new plugin project
2. Building a DLL
3. Copying it into `/Modules`

No host application modification should be required.

---

# 22. Non-Goals

This project is not intended to:

* replace enterprise systems
* integrate external web services
* act as a large distributed platform

It is strictly an **internal productivity tool platform**.

---

# 23. Core Principles

All development decisions should prioritize:

1. portability
2. modularity
3. reliability
4. offline capability
5. enterprise safety

Simplicity is preferred over complexity.
