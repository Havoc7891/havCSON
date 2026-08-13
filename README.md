# havCSON

Havoc's single-file CSON (CoffeeScript Object Notation) library for C++.

## Table of Contents

- [Features](#features)
- [Getting Started](#getting-started)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Features

- Single-header, dependency-free (just include `havCSON.hpp`)
- Reading and writing CSON files (from strings or UTF-8 file paths)
  - Also supports generating CSON files from scratch
  - Atomic file writers safely replace an existing file after the new contents are flushed
- Error-code and optional throwing parse APIs, including lossless parsing helpers
- Duplicate object keys are rejected with `ErrorCode::DuplicateKey`
- Finite `double` values use round-trip-safe parsing and formatting
- Convert parsed data to JSON text via `ToJsonString`
- Pretty-print output with controllable indent width and optional key sorting
- Comment-aware round trips preserve comments (including comments after `:` or before a closing delimiter), blank lines, and member ordering while regenerating normalized CSON
- Unicode / UTF-8 support with validation
- Compile-time version constants (`VersionMajor`, `VersionMinor`, `VersionPatch`, and `VersionString`)

## Getting Started

This library requires C++23. The library should be cross-platform, but has only been tested under Windows so far.

### Installation

Copy the header file into your project folder and include the file like this:

```cpp
#include "havCSON.hpp"
```

### Usage

Here are some code examples demonstrating how to use the library. The iteration and lookup snippets reuse `player` from the array-construction example; the other snippets are independent.

#### Read CSON file

```cpp
using namespace havCSON;

Value config;
Error error;

auto code = ParseFile("config.cson", config, &error);
if (code != ErrorCode::OK)
{
  std::cerr << "Failed to parse " << error.filename << ": line " << error.where.line
            << " col " << error.where.column
            << " (" << error.message << ")\n";
  return;
}

// Optional: throw on error
Value parsed = ParseOrThrow("foo: 1");

// Convert to JSON string without pretty-printing
std::string json;
if (!ToJsonString(config, json, &error))
{
  std::cerr << "JSON conversion failed: " << error.message << "\n";
}
```

#### Reject duplicate object keys

```cpp
using namespace havCSON;

Value value;
Error error;
auto code = Parse("name: \"first\"\nname: \"second\"\n", value, &error);
if (code == ErrorCode::DuplicateKey)
{
  std::cerr << error.message << " at line " << error.where.line << "\n";
}
```

Duplicate keys are rejected in both normal and lossless parsing instead of silently replacing or ignoring a value.

#### Write CSON file atomically

Numbers are stored as `double`. Finite values are formatted so that writing and parsing them recovers the same value; the original numeric spelling is not retained.

```cpp
using namespace havCSON;

Value root = Object{
  {"name", "Havoc"},
  {"level", 42.0}, // Numbers are double
  {"items", Array{1.0, 2.0, 3.0}},
};

WriteOptions options;
options.indentWidth = 2;
options.sortObjectKeys = true;

Error error;
if (!WriteFileAtomic("out.cson", root, options, &error))
{
  std::cerr << "Atomic write failed: " << error.message << "\n";
}

// Get the formatted string without touching disk
std::string text;
if (!ToString(root, text, options, &error))
{
  std::cerr << "Formatting failed: " << error.message << "\n";
}
```

`WriteFileAtomic` serializes to a temporary file in the destination directory, flushes it, and then replaces the destination. Use `WriteFile` for a direct write, `WriteTextFileAtomic` for text that is already serialized, or `WriteFileLosslessAtomic` for a `LosslessValue`.

#### Parse from a string and mutate the data

```cpp
using namespace havCSON;

// CoffeeScript-style inline array and object
Value value = ParseOrThrow("user: {name: \"Havoc\", tags: [\"dev\", \"cpp\"]}");

// Add a nested field
Object& obj = value.asObject().at("user").asObject();
obj["active"] = true;

// Push another tag
obj["tags"].asArray().push_back("oss");

Error error;
std::string formatted;
if (!ToString(value, formatted, {}, &error))
{
  std::cerr << "Formatting failed: " << error.message << "\n";
}
else
{
  std::cout << formatted << "\n";
}
```

#### Create an array and array entries

```cpp
using namespace havCSON;

Array inventory;
inventory.emplace_back("Sword");
inventory.emplace_back("Shield");
inventory.emplace_back(3.0);

Value player = Object{
  {"name", "Havoc"},
  {"inventory", inventory},
};
```

#### Iterate over an array

```cpp
const havCSON::Array& items = player.asObject().at("inventory").asArray();
for (const havCSON::Value& item : items)
{
  std::cout << "Item type index: " << item.index() << "\n";
}
```

#### Check if value exists

```cpp
const auto& obj = player.asObject();
if (auto it = obj.find("name"); it != obj.end())
{
  std::cout << "Name: " << std::get<std::string>(it->second) << "\n";
}
```

#### Check value type

```cpp
const havCSON::Value& numberValue = player.asObject().at("inventory").asArray().at(2);
if (numberValue.isNumber())
{
  double level = std::get<double>(numberValue);
  std::cout << "Level: " << level << "\n";
}
```

#### Handle multiline strings and comments losslessly

```cpp
using namespace havCSON;

std::string src = R"(# Header comment
bio: """
  Multiline
  text preserved
"""
)";

LosslessValue lossless;
Error error;
if (ParseLossless(src, lossless, &error) != ErrorCode::OK)
{
  // Handle
}

// Preserve comments, blank lines, and ordering while normalizing formatting
std::string roundtrip;
if (!ToStringLossless(lossless, roundtrip, {}, &error))
{
  std::cerr << "Lossless formatting failed: " << error.message << "\n";
}

if (!WriteFileLosslessAtomic("out.cson", lossless, {}, &error))
{
  std::cerr << "Lossless atomic write failed: " << error.message << "\n";
}
```

Use `ParseFileLossless` to read directly from a UTF-8 file. `ParseLosslessOrThrow` is the throwing counterpart for string input. Lossless mode retains blank lines, member ordering, and comments in structural positions such as after `key:` and before closing delimiters. It regenerates normalized formatting rather than preserving the original bytes.

## Contributing

Thank you for your interest! Suggestions for features and bug reports are always welcome via issues.

To maintain a consistent design and quality for this library, changes are implemented by the maintainer rather than via direct pull requests.

## License

Copyright &copy; 2025-2026 Ren&eacute; Nicolaus

This library is released under the [MIT license](/LICENSE).
