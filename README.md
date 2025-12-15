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
- Optional throwing parse API (`ParseOrThrow`) in addition to error-code based parsing
- Convert parsed data to JSON text via `ToJsonString`
- Pretty-print output with controllable indent width and optional key sorting
- Lossless round-trip mode that keeps comments / blank lines / ordering with matching write helpers
- Unicode / UTF-8 support with validation

## Getting Started

This library requires C++23. The library should be cross-platform, but has only been tested under Windows so far.

### Installation

Copy the header file into your project folder and include the file like this:

```cpp
#include "havCSON.hpp"
```

### Usage

Here are some code examples demonstrating how to use the library. Each snippet is independent; copy the header and drop the snippet you need into your code.

#### Read CSON file

```cpp
using namespace havCSON;

Value config;
Error error;

auto code = ParseFile("config.cson", config, &error);
if (code != ErrorCode::OK)
{
  std::cerr << "Failed: line " << error.where.line << " col " << error.where.column
            << " (" << error.message << ")\n";
  return;
}

// Optional: throw on error
Value parsed = ParseOrThrow("foo: 1");

// Convert to JSON string without pretty-printing
std::string json = ToJsonString(config);
```

#### Write CSON file

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
if (!WriteFile("out.cson", root, options, &error))
{
  std::cerr << "Write failed: " << error.message << "\n";
}

// Get the formatted string without touching disk
std::string text = ToString(root, options);
```

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

std::cout << ToString(value) << "\n";
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

// Keep comments / spacing when writing back out
std::string roundtrip = ToStringLossless(lossless);
```

## Contributing

Feel free to suggest features or report issues. However, please note that pull requests will not be accepted.

## License

Copyright &copy; 2025 Ren&eacute; Nicolaus

This library is released under the [MIT license](/LICENSE).
