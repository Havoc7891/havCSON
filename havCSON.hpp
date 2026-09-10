/*
havCSON.hpp

ABOUT

Havoc's single-file CSON (CoffeeScript Object Notation) library for C++.

REVISION HISTORY

v0.5.0 (2026-09-10)
- Added optional source locations, safe lossless editing, parser resource limits, structured file errors, and source-aware checked access.
- Fixed parser state resets, identifier rewinding, newline handling, and lossless round trips.

v0.4.0 (2026-08-13)
- Added duplicate-key rejection, made finite double parsing and formatting round-trip-safe, fixed lossless parsing and writing, and added atomic file writers.

v0.3.0 (2026-05-15)
- Simplified platform type size checks.

v0.2.0 (2026-01-18)
- Trimmed float output, non-finite numbers are rejected on write, added error-returning writer overloads.

v0.1.0 (2025-12-15)
- First release.

LICENSE

MIT License

Copyright (c) 2025-2026 René Nicolaus

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

#ifndef HAVCSON_HPP
#define HAVCSON_HPP

#ifdef _WIN32
  #ifdef _MBCS
    #error "_MBCS is defined, but only Unicode is supported"
  #endif
  #undef _UNICODE
  #define _UNICODE 1
  #undef UNICODE
  #define UNICODE 1

  #undef NOMINMAX
  #define NOMINMAX

  #undef STRICT
  #define STRICT

  #ifndef _WIN32_WINNT
    // Current MinGW builds use Windows 7 APIs for C++23 threading support
    #define _WIN32_WINNT 0x0601
  #endif
  #ifdef _MSC_VER
    #include <SDKDDKVer.h>
  #endif

  #undef WIN32_LEAN_AND_MEAN
  #define WIN32_LEAN_AND_MEAN
  #include <windows.h>
#endif

#include <cstdio>
#include <cstdint>
#include <cstdlib>
#include <climits>
#include <array>
#include <algorithm>
#include <charconv>
#include <cmath>
#include <cerrno>
#include <exception>
#include <expected>
#include <concepts>
#include <functional>
#include <fstream>
#include <limits>
#include <memory>
#include <optional>
#include <sstream>
#include <string>
#include <string_view>
#include <system_error>
#include <type_traits>
#include <unordered_map>
#include <utility>
#include <variant>
#include <vector>

#ifndef _WIN32
  #include <atomic>
  #include <fcntl.h>
  #include <sys/stat.h>
  #include <sys/types.h>
  #include <unistd.h>
#endif

static_assert(CHAR_BIT == 8, "havCSON requires 8-bit bytes");

namespace havCSON
{
  inline constexpr std::uint32_t VersionMajor = 0;
  inline constexpr std::uint32_t VersionMinor = 5;
  inline constexpr std::uint32_t VersionPatch = 0;
  inline constexpr std::string_view VersionString = "0.5.0";

  struct FileCloser
  {
    void operator()(std::FILE* file) const noexcept
    {
      if (file)
      {
        std::fclose(file);
      }
    }
  };

  using FilePtr = std::unique_ptr<std::FILE, FileCloser>;

#ifdef _WIN32
  // Convert UTF-8 to UTF-16; keep the trailing null when forFileStream is true so _wfopen / _wifstream can use data()
  inline std::wstring ConvertStringToWString(const std::string& value, bool forFileStream = false)
  {
    int numChars = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, value.c_str(), -1, nullptr, 0);
    if (numChars <= 0)
    {
      return {};
    }

    std::wstring wstr(static_cast<std::size_t>(numChars), L'\0');
    int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, value.c_str(), -1, wstr.data(), numChars);
    if (written <= 0)
    {
      return {};
    }

    // For general string use, strip the extra null terminator added by MultiByteToWideChar
    if (!forFileStream && !wstr.empty() && wstr.back() == L'\0')
    {
      wstr.pop_back();
    }

    return wstr;
  }

  // Cross-platform FILE opener that accepts UTF-8 paths and uses wide APIs on Windows
  inline FilePtr OpenFileUTF8(const std::string& path, const std::string& mode)
  {
    FilePtr fileStream;
    if (path.find('\0') != std::string::npos || mode.find('\0') != std::string::npos)
    {
      errno = EINVAL;
      return fileStream;
    }
    std::wstring modeW = ConvertStringToWString(mode, true);
    std::wstring pathW = ConvertStringToWString(path, true);
    if (modeW.empty() || pathW.empty())
    {
      errno = EILSEQ;
      return fileStream;
    }
    std::FILE* file = nullptr;
    const auto result = _wfopen_s(&file, pathW.c_str(), modeW.c_str());
    if (result == 0)
    {
      fileStream.reset(file);
    }
    else
    {
      errno = result;
    }
    return fileStream;
  }
#else
  inline FilePtr OpenFileUTF8(const std::string& path, const std::string& mode)
  {
    FilePtr fileStream;
    if (path.find('\0') != std::string::npos || mode.find('\0') != std::string::npos)
    {
      errno = EINVAL;
      return fileStream;
    }
    fileStream.reset(std::fopen(path.c_str(), mode.c_str()));
    return fileStream;
  }
#endif

  struct LocationEntry
  {
    std::size_t line = 1;
    std::size_t column = 1;
  };

  struct ParseOptions
  {
    bool trackSourceLocations = false;
    // Container nesting includes the root object/array. Scalars have depth zero.
    // Zero disables either limit. The input limit counts original UTF-8 bytes,
    // including a BOM, and is checked before validation or source indexing.
    std::size_t maxDepth = 256;
    std::size_t maxInputBytes = 0;
  };

  // Offsets address the original UTF-8 bytes (including a BOM). Lines and byte
  // columns are one-based. LF, CRLF, and CR each start a new line.
  struct SourcePosition
  {
    std::size_t byteOffset = 0;
    std::size_t line = 1;
    std::size_t column = 1;

    friend bool operator==(const SourcePosition&, const SourcePosition&) = default;
  };

  struct SourceSpan
  {
    SourcePosition begin;
    SourcePosition end; // Exclusive

    friend bool operator==(const SourceSpan&, const SourceSpan&) = default;
  };

  struct SourceInfo
  {
    SourceSpan valueSpan;
    std::optional<SourceSpan> keySpan; // Present only for an object member
    std::shared_ptr<const std::string> filename; // Owned - null for unnamed input
  };

  enum class ErrorCode : std::uint8_t
  {
    OK,
    UnexpectedChar,
    UnexpectedEnd,
    InvalidNumber,
    InvalidEscape,
    InvalidUtf8,
    UnterminatedString,
    UnterminatedTripleString,
    InvalidIndentChar,
    InconsistentIndent,
    InternalError,
    DuplicateKey,
    ResourceLimit,
    IoError,
    InvalidPath,
    TypeMismatch,
    InvalidLosslessTree,
    OutOfRange,
    MissingMember,
  };

  struct Error
  {
    ErrorCode code = ErrorCode::OK;
    LocationEntry where{};
    std::string message;
    std::string filename;
    // File failures retain their operation and native cause before cleanup.
    // These remain empty for syntax, validation, and resource-limit errors.
    std::string operation;
    std::error_code systemError;

    explicit operator bool() const
    {
      return code != ErrorCode::OK;
    }
  };

  struct Value;

  using Array = std::vector<Value>;
  using Object = std::unordered_map<std::string, Value>;

  struct Value : std::variant<std::nullptr_t, bool, double, std::string, Array, Object>
  {
    using variant::variant;

    // Original source locations, not live positions in regenerated output.
    // Copies retain them. Newly constructed values have none. Metadata is not
    // part of semantic comparisons and is never serialized by a writer.
    const SourceInfo* Source() const noexcept
    {
      return mSource.get();
    }

    void ClearSource() noexcept
    {
      mSource.reset();
    }

    bool isNull() const
    {
      return std::holds_alternative<std::nullptr_t>(*this);
    }

    bool isBool() const
    {
      return std::holds_alternative<bool>(*this);
    }

    bool isNumber() const
    {
      return std::holds_alternative<double>(*this);
    }

    bool isString() const
    {
      return std::holds_alternative<std::string>(*this);
    }

    bool isArray() const
    {
      return std::holds_alternative<Array>(*this);
    }

    bool isObject() const
    {
      return std::holds_alternative<Object>(*this);
    }

    const Array& asArray() const
    {
      return std::get<Array>(*this);
    }

    const Object& asObject() const
    {
      return std::get<Object>(*this);
    }

    Array& asArray()
    {
      return std::get<Array>(*this);
    }

    Object& asObject()
    {
      return std::get<Object>(*this);
    }

  private:
    friend class Parser;
    std::shared_ptr<SourceInfo> mSource;
  };

  // Optional lossless representation that can carry comments / ordering for regeneration
  struct LosslessComment
  {
    int indent = 0; // Indent columns where this comment line began
    std::string text; // Comment text without trailing newline (empty -> blank line)
  };

  struct LosslessValue
  {
    Value value;
    std::vector<LosslessComment> leadingComments; // Full lines with recorded indent
    std::string inlineComment; // Text after '#' on the same line as the value
    std::vector<LosslessValue> arrayItems; // In-order children if value is array
    std::vector<std::pair<std::string, LosslessValue>> objectItems; // In-order children if value is object
    std::vector<LosslessComment> trailingComments; // Comments / blank lines after this value (before dedent)
    std::string blockComment; // Comment after an object key's ':' before a block value
    std::vector<LosslessComment> closingComments; // Full-line comments immediately before a container's closing delimiter

    const SourceInfo* Source() const noexcept
    {
      return value.Source();
    }
  };

  inline void ClearSourceLocations(Value& value) noexcept
  {
    value.ClearSource();

    if (value.isArray())
    {
      for (auto& child : value.asArray())
      {
        ClearSourceLocations(child);
      }
    }
    else if (value.isObject())
    {
      for (auto& member : value.asObject())
      {
        ClearSourceLocations(member.second);
      }
    }
  }

  inline void ClearSourceLocations(LosslessValue& value) noexcept
  {
    ClearSourceLocations(value.value);

    for (auto& child : value.arrayItems)
    {
      ClearSourceLocations(child);
    }

    for (auto& member : value.objectItems)
    {
      ClearSourceLocations(member.second);
    }
  }

  // Paths own their keys. A string is always one literal object key, never a
  // dot-separated expression. This also supports empty keys and keys with dots.
  using ValuePathSegment = std::variant<std::string, std::size_t>;
  using ValuePath = std::vector<ValuePathSegment>;

  inline std::string FormatValuePath(const ValuePath& path)
  {
    std::string result = "$";
    constexpr char hexDigits[] = "0123456789abcdef";

    for (const auto& segment : path)
    {
      if (const auto* key = std::get_if<std::string>(&segment))
      {
        result += "[\"";

        for (const unsigned char byte : *key)
        {
          switch (byte)
          {
            case '"': result += "\\\""; break;
            case '\\': result += "\\\\"; break;
            case '\b': result += "\\b"; break;
            case '\f': result += "\\f"; break;
            case '\n': result += "\\n"; break;
            case '\r': result += "\\r"; break;
            case '\t': result += "\\t"; break;
            default:
              if (byte < 0x20)
              {
                result += "\\u00";
                result += hexDigits[byte >> 4];
                result += hexDigits[byte & 0x0f];
              }
              else
              {
                result += static_cast<char>(byte);
              }
              break;
          }
        }

        result += "\"]";
      }
      else
      {
        result += '[';
        result += std::to_string(std::get<std::size_t>(segment));
        result += ']';
      }
    }

    return result;
  }

  inline std::string_view ValueTypeName(const Value& value) noexcept
  {
    switch (value.index())
    {
      case 0: return "null";
      case 1: return "boolean";
      case 2: return "number";
      case 3: return "string";
      case 4: return "array";
      case 5: return "object";
      default: return "invalid value";
    }
  }

  // Unlike a ValueView, this diagnostic owns all of its data. Source locations
  // are copied, so the error remains useful after the document is destroyed.
  struct AccessError
  {
    ErrorCode code = ErrorCode::TypeMismatch;
    ValuePath path;
    std::string message;
    std::string expectedType;
    std::string actualType;
    std::optional<SourceInfo> source;
    // Missing members and out-of-range indices refer to their parent value.
    // Such errors never claim to know the source span of the missing child.
    bool sourceIsContainer = false;
  };

  // A borrowed read-only cursor. The document must outlive the cursor and any
  // reference returned by AsString/AsArray/AsObject/Get, and must not be mutated
  // while these are used. Navigation retains the first error for easy chaining.
  class ValueView
  {
  public:
    explicit ValueView(const Value& value) noexcept : mValue(&value) {}
    ValueView(Value&&) = delete;
    ValueView(const Value&&) = delete;

    const ValuePath& Path() const noexcept
    {
      return mPath;
    }

    bool HasValue() const noexcept
    {
      return !mError.has_value();
    }

    std::expected<std::reference_wrapper<const Value>, AccessError> Get() const
    {
      if (mError)
      {
        return std::unexpected(*mError);
      }
      return std::cref(*mValue);
    }

    ValueView Member(std::string_view key) const
    {
      if (mError)
      {
        return *this;
      }

      const auto* object = std::get_if<Object>(mValue);
      if (!object)
      {
        return ValueView(TypeError("object"));
      }

      ValuePath path = mPath;
      path.emplace_back(std::string(key));

      const auto found = object->find(std::string(key));
      if (found == object->end())
      {
        return ValueView(MakeError(ErrorCode::MissingMember, std::move(path),
            "member", "missing", "required member is missing", true));
      }

      return ValueView(found->second, std::move(path));
    }

    ValueView At(std::size_t index) const
    {
      if (mError)
      {
        return *this;
      }

      const auto* array = std::get_if<Array>(mValue);
      if (!array)
      {
        return ValueView(TypeError("array"));
      }

      ValuePath path = mPath;
      path.emplace_back(index);

      if (index >= array->size())
      {
        return ValueView(MakeError(ErrorCode::OutOfRange, std::move(path),
            "array element", "missing", "array index is out of range (size " +
            std::to_string(array->size()) + ')', true));
      }

      return ValueView((*array)[index], std::move(path));
    }

    // Resolve literal segments relative to this cursor - an empty path is a no-op
    ValueView AtPath(const ValuePath& path) const
    {
      ValueView view = *this;

      for (const auto& segment : path)
      {
        if (const auto* key = std::get_if<std::string>(&segment))
        {
          view = view.Member(*key);
        }
        else
        {
          view = view.At(std::get<std::size_t>(segment));
        }

        if (!view.HasValue())
        {
          break;
        }
      }

      return view;
    }

    std::expected<std::nullptr_t, AccessError> AsNull() const
    {
      return ReadScalar<std::nullptr_t>("null");
    }

    std::expected<bool, AccessError> AsBool() const
    {
      return ReadScalar<bool>("boolean");
    }

    std::expected<double, AccessError> AsNumber() const
    {
      return ReadScalar<double>("number");
    }

    std::expected<std::reference_wrapper<const std::string>, AccessError> AsString() const
    {
      return ReadReference<std::string>("string");
    }

    std::expected<std::reference_wrapper<const Array>, AccessError> AsArray() const
    {
      return ReadReference<Array>("array");
    }

    std::expected<std::reference_wrapper<const Object>, AccessError> AsObject() const
    {
      return ReadReference<Object>("object");
    }

    // Numbers still have double semantics: this checks the stored number, not
    // precision already lost before parsing. Store exact large IDs as strings.
    template <std::integral Integer>
      requires (!std::same_as<std::remove_cv_t<Integer>, bool> &&
                std::same_as<Integer, std::remove_cv_t<Integer>>)
    std::expected<Integer, AccessError> AsInteger() const
    {
      const auto number = AsNumber();

      if (!number)
      {
        if (mError)
        {
          return std::unexpected(number.error());
        }

        return std::unexpected(TypeError("integer"));
      }

      if (!std::isfinite(*number) || std::trunc(*number) != *number)
      {
        return std::unexpected(MakeError(ErrorCode::InvalidNumber, mPath,
            "integer", "number", "expected a finite integral number"));
      }

      // The upper bound is exclusive. Converting INT64_MAX or UINT64_MAX to
      // double rounds them up, so a <= double(max) check would allow UB here.
      constexpr int digits = std::numeric_limits<Integer>::digits;
      const double upper = std::ldexp(1.0, digits);
      const double lower = std::is_signed_v<Integer> ? -upper : 0.0;

      if (*number < lower || *number >= upper)
      {
        return std::unexpected(MakeError(ErrorCode::OutOfRange, mPath,
            "integer", "number", "number is outside the requested integer type's range"));
      }

      return static_cast<Integer>(*number);
    }

    template <std::integral Integer>
      requires (!std::same_as<std::remove_cv_t<Integer>, bool> &&
                std::same_as<Integer, std::remove_cv_t<Integer>>)
    std::expected<Integer, AccessError> AsInteger(Integer minimum, Integer maximum) const
    {
      const auto integer = AsInteger<Integer>();

      if (!integer)
      {
        return integer;
      }

      if (minimum > maximum || *integer < minimum || *integer > maximum)
      {
        return std::unexpected(MakeError(ErrorCode::OutOfRange, mPath,
            "integer", "number", "number is outside the required range [" +
            std::to_string(minimum) + ", " + std::to_string(maximum) + ']'));
      }

      return integer;
    }

  private:
    ValueView(const Value& value, ValuePath path)
      : mValue(&value), mPath(std::move(path)) {}

    explicit ValueView(AccessError error)
      : mValue(nullptr), mPath(error.path), mError(std::move(error)) {}

    AccessError MakeError(ErrorCode code, ValuePath path, std::string expected,
                          std::string actual, std::string detail,
                          bool sourceIsContainer = false) const
    {
      AccessError error;
      error.code = code;
      error.path = std::move(path);
      error.expectedType = std::move(expected);
      error.actualType = std::move(actual);
      error.message = FormatValuePath(error.path) + ": " + std::move(detail);
      error.sourceIsContainer = sourceIsContainer;

      if (mValue && mValue->Source())
      {
        error.source = *mValue->Source();

        if (sourceIsContainer)
        {
          error.source->keySpan.reset();
        }
      }

      return error;
    }

    AccessError TypeError(std::string expected) const
    {
      const std::string actual(ValueTypeName(*mValue));
      const std::string detail = "expected " + expected + ", got " + actual;

      return MakeError(ErrorCode::TypeMismatch, mPath, std::move(expected), actual, detail);
    }

    template <typename Type>
    std::expected<Type, AccessError> ReadScalar(std::string_view expected) const
    {
      if (mError)
      {
        return std::unexpected(*mError);
      }

      if (const auto* value = std::get_if<Type>(mValue))
      {
        return *value;
      }

      return std::unexpected(TypeError(std::string(expected)));
    }

    template <typename Type>
    std::expected<std::reference_wrapper<const Type>, AccessError>
    ReadReference(std::string_view expected) const
    {
      if (mError)
      {
        return std::unexpected(*mError);
      }

      if (const auto* value = std::get_if<Type>(mValue))
      {
        return std::cref(*value);
      }

      return std::unexpected(TypeError(std::string(expected)));
    }

    const Value* mValue;
    ValuePath mPath;
    std::optional<AccessError> mError;
  };

  class Parser
  {
  public:
    Parser(std::string_view src, std::string_view filename = {}, const ParseOptions& options = {})
      : mSrc(src), mFilename(filename), mOptions(options)
    {
      if (mOptions.trackSourceLocations && !InputLimitExceeded())
      {
        if (!mFilename.empty())
        {
          mSourceFilename = std::make_shared<const std::string>(mFilename);
        }

        const bool bom = mSrc.starts_with("\xEF\xBB\xBF");

        mSourceLineStarts.push_back(bom ? 3 : 0);

        for (std::size_t offset = 0; offset < mSrc.size(); ++offset)
        {
          if (mSrc[offset] == '\r')
          {
            if (offset + 1 < mSrc.size() && mSrc[offset + 1] == '\n')
            {
              ++offset;
            }

            mSourceLineStarts.push_back(offset + 1);
          }
          else if (mSrc[offset] == '\n')
          {
            mSourceLineStarts.push_back(offset + 1);
          }
        }
      }
    }

    ErrorCode Parse(Value& out, Error* error = nullptr)
    {
      ResetParseState();

      if (InputLimitExceeded())
      {
        return Fail(ErrorCode::ResourceLimit, error, "Maximum input byte count exceeded");
      }

      // Validate UTF-8 up front (strips leading BOM)
      std::size_t badIndex = 0;
      std::size_t badLine = 1;
      std::size_t badCol = 1;
      bool hasBOM = mSrc.size() >= 3 && static_cast<unsigned char>(mSrc[0]) == 0xEF && static_cast<unsigned char>(mSrc[1]) == 0xBB &&
        static_cast<unsigned char>(mSrc[2]) == 0xBF;
      if (!ValidateUTF8(mSrc, true, badIndex, badLine, badCol))
      {
        mPos = badIndex;
        mLine = badLine;
        mCol = badCol;
        return Fail(ErrorCode::InvalidUtf8, error, "Invalid UTF-8 encoding");
      }
      if (hasBOM)
      {
        mPos = 3;
        mCol = 1;
        mLine = 1;
      }

      Value value;
      SkipWhitespaceAndComments();
      if (mPos >= mSrc.size())
      {
        // Empty document -> null
        value = nullptr;
      }
      else
      {
        ErrorCode errorCode = ParseValue(value, 0);
        if (errorCode != ErrorCode::OK)
        {
          if (error)
          {
            if (mError.code != ErrorCode::OK)
            {
              *error = mError;
            }
            else
            {
              *error = {};
              error->code = errorCode;
              error->where = Location();
              error->message.clear();
              error->filename = mFilename;
            }
          }
          return errorCode;
        }
      }

      // After parsing top-level value, consume any remaining whitespace / comments and ensure we're back at indent
      // level 0
      while (!EndOfFile())
      {
        char c = Peek();
        if (c == ' ' || c == '\t' || c == '\r' || c == '\n')
        {
          Get();
          continue;
        }
        if (c == '#')
        {
          SkipToEOL();
          continue;
        }
        break;
      }

      // Reset indent stack to base level
      mIndentStack.clear();
      mIndentStack.push_back(0);

      if (mPos != mSrc.size())
      {
        return Fail(ErrorCode::UnexpectedChar, error, "Trailing characters after top-level value");
      }
      out = std::move(value);
      if (error)
      {
        *error = {};
      }
      return ErrorCode::OK;
    }

    const Error& LastError() const
    {
      return mError;
    }

  protected:
    std::string_view mSrc;
    std::string mFilename;
    std::size_t mPos = 0;
    std::size_t mLine = 1;
    std::size_t mCol = 1;

    // CoffeeScript-like indent model
    int mIndentUnit = 0; // Discovered on first non-zero indent
    std::vector<int> mIndentStack{0}; // Known indent levels (columns)

    Error mError;
    ParseOptions mOptions;
    std::shared_ptr<const std::string> mSourceFilename;
    std::vector<std::size_t> mSourceLineStarts;
    std::optional<std::size_t> mArrayClosingEnd;
    std::size_t mContainerDepth = 0;

    struct ContainerDepthGuard
    {
      explicit ContainerDepthGuard(std::size_t& depth) : mDepth(depth) { ++mDepth; }
      ~ContainerDepthGuard() { --mDepth; }
      ContainerDepthGuard(const ContainerDepthGuard&) = delete;
      ContainerDepthGuard& operator=(const ContainerDepthGuard&) = delete;

    private:
      std::size_t& mDepth;
    };

    bool InputLimitExceeded() const noexcept
    {
      return mOptions.maxInputBytes != 0 && mSrc.size() > mOptions.maxInputBytes;
    }

    bool CanEnterContainer()
    {
      if (mOptions.maxDepth != 0 && mContainerDepth >= mOptions.maxDepth)
      {
        Fail(ErrorCode::ResourceLimit, nullptr, "Maximum container nesting depth exceeded");
        return false;
      }

      return true;
    }

    void ResetParseState()
    {
      mError = {};
      mPos = 0;
      mLine = mCol = 1;
      mIndentUnit = 0;
      mIndentStack.assign(1, 0);
      mArrayClosingEnd.reset();
      mContainerDepth = 0;
    }

    SourcePosition SourcePositionAt(std::size_t offset) const
    {
      if (!mOptions.trackSourceLocations)
      {
        return {};
      }

      const auto next = std::upper_bound(mSourceLineStarts.begin(), mSourceLineStarts.end(), offset);
      const auto index = next == mSourceLineStarts.begin() ? 0 :
        static_cast<std::size_t>(next - mSourceLineStarts.begin() - 1);
      const auto start = mSourceLineStarts[index];

      return {offset, index + 1, offset >= start ? offset - start + 1 : 1};
    }

    void RecordSource(Value& value, std::size_t begin)
    {
      if (!mOptions.trackSourceLocations)
      {
        return;
      }

      std::size_t end = mPos;

      if ((value.isObject() && mSrc[begin] != '{') ||
          (value.isArray() && !mArrayClosingEnd))
      {
        // Indentation parsing may already have consumed following comments or
        // the next sibling's indentation. The final child, not the cursor,
        // determines the end of a block value.
        end = begin;
        const auto include = [&end](const Value& child)
        {
          if (const auto* source = child.Source())
          {
            end = std::max(end, source->valueSpan.end.byteOffset);
          }
        };

        if (value.isObject())
        {
          for (const auto& member : value.asObject())
          {
            include(member.second);
          }
        }
        else
        {
          end = begin + 1; // Opening '[' of an empty multiline array
          for (const auto& child : value.asArray())
          {
            include(child);
          }
        }
      }
      else if (value.isArray())
      {
        end = *mArrayClosingEnd;
      }
      else if (!value.isObject())
      {
        // Identifier/string lookahead consumes spaces while looking for ':'.
        // Quoted literals end with their quote, so stripping exterior space
        // never removes bytes belonging to their decoded string contents.
        while (end > begin && (mSrc[end - 1] == ' ' || mSrc[end - 1] == '\t' ||
                               mSrc[end - 1] == '\r' || mSrc[end - 1] == '\n'))
        {
          --end;
        }
      }

      value.mSource = std::make_shared<SourceInfo>(SourceInfo{
        {SourcePositionAt(begin), SourcePositionAt(end)}, std::nullopt, mSourceFilename});
    }

    void RecordKeySource(Value& value, const SourceSpan& key)
    {
      if (!mOptions.trackSourceLocations || !value.mSource)
      {
        return;
      }

      if (!value.mSource.unique())
      {
        value.mSource = std::make_shared<SourceInfo>(*value.mSource);
      }

      value.mSource->keySpan = key;
    }

    char Peek() const
    {
      return mPos < mSrc.size() ? mSrc[mPos] : '\0';
    }

    bool EndOfFile() const
    {
      return mPos >= mSrc.size();
    }

    char Get()
    {
      if (mPos >= mSrc.size())
      {
        return '\0';
      }
      char c = mSrc[mPos++];
      if (c == '\r' || (c == '\n' && (mPos < 2 || mSrc[mPos - 2] != '\r')))
      {
        ++mLine;
        mCol = 1;
      }
      else if (c != '\n')
      {
        ++mCol;
      }
      return c;
    }

    bool Match(char c)
    {
      if (Peek() == c)
      {
        Get();
        return true;
      }
      return false;
    }

    LocationEntry Location() const
    {
      return LocationEntry{mLine, mCol};
    }

    ErrorCode Fail(ErrorCode code, Error* out, std::string_view message = {})
    {
      mError.code = code;
      mError.where = Location();
      mError.message.assign(message.begin(), message.end());
      mError.filename.assign(mFilename.begin(), mFilename.end());
      if (out)
      {
        *out = mError;
      }
      return code;
    }

    ErrorCode FailAt(ErrorCode code, LocationEntry where, std::string_view message = {})
    {
      mError.code = code;
      mError.where = where;
      mError.message.assign(message.begin(), message.end());
      mError.filename.assign(mFilename.begin(), mFilename.end());
      return code;
    }

    ErrorCode CheckDuplicateKey(const Object& object, const std::string& key, LocationEntry where)
    {
      if (object.find(key) == object.end())
      {
        return ErrorCode::OK;
      }
      return FailAt(ErrorCode::DuplicateKey, where, "Duplicate object key '" + key + "'");
    }

    void SkipInlineSpaces()
    {
      while (true)
      {
        char c = Peek();
        if (c == ' ' || c == '	')
        {
          Get();
        }
        else
        {
          break;
        }
      }
    }

    void SkipToEOL()
    {
      while (!EndOfFile())
      {
        char c = Get();
        if (c == '\r' || c == '\n')
        {
          if (c == '\r' && Peek() == '\n')
          {
            Get();
          }
          break;
        }
      }
    }

    void SkipWhitespaceAndComments()
    {
      while (!EndOfFile())
      {
        char c = Peek();
        if (c == ' ' || c == '	')
        {
          Get();
          continue;
        }
        if (c == '#')
        {
          SkipToEOL();
          continue;
        }
        if (c == '\r')
        {
          Get();
          if (Peek() == '\n')
          {
            Get();
          }
          continue;
        }
        if (c == '\n')
        {
          Get();
          continue;
        }
        break;
      }
    }

    // Reads indentation (spaces only) at the beginning of a line.
    // Assumes we are positioned at the first character of a line.
    // - indentCols: number of columns (spaces).
    // - hasContent: true if the line has non-comment content.
    // Returns ErrorCode::Ok or an indent-related error.
    ErrorCode ReadLineIndent(int& indentCols, bool& hasContent)
    {
      indentCols = 0;
      hasContent = false;

      // Consume leading spaces / detect tabs in indent
      while (!EndOfFile())
      {
        char c = Peek();
        if (c == ' ')
        {
          Get();
          ++indentCols;
        }
        else if (c == '\t')
        {
          // Tabs are illegal in indent
          return Fail(ErrorCode::InvalidIndentChar, nullptr, "Tabs are not allowed in indentation");
        }
        else
        {
          break;
        }
      }

      // Check if the rest of the line is blank or comment-only
      char c = Peek();
      if (c == '#' || c == '\r' || c == '\n' || c == '\0')
      {
        // Skip comment content
        if (c == '#')
        {
          SkipToEOL();
          // SkipToEOL already consumed this comment's line ending. Leave a
          // following blank line for the next call so lossless parsing can
          // retain it as a separate trivia item.
          return ErrorCode::OK;
        }
        // Consume line ending (if any)
        if (c == '\r')
        {
          Get();
          if (Peek() == '\n')
          {
            Get();
          }
        }
        else if (c == '\n')
        {
          Get();
        }
        // No content on this line
        hasContent = false;
        return ErrorCode::OK;
      }

      // This line has content
      hasContent = true;

      // Enforce a global indent unit
      if (indentCols > 0)
      {
        if (mIndentUnit == 0)
        {
          mIndentUnit = indentCols;
        }
        else if (indentCols % mIndentUnit != 0)
        {
          return Fail(ErrorCode::InconsistentIndent, nullptr, "Indentation is not a multiple of base indent width");
        }
      }

      return ErrorCode::OK;
    }

    // Apply CoffeeScript-style indent stack transitions:
    // - If indent > top: push (new block).
    // - If indent < top: pop until match or error.
    // - If indent == top: stay in current block.
    ErrorCode ApplyIndentLevel(int indentCols)
    {
      int top = mIndentStack.back();
      if (indentCols > top)
      {
        mIndentStack.push_back(indentCols);
      }
      else if (indentCols < top)
      {
        while (!mIndentStack.empty() && indentCols < mIndentStack.back())
        {
          mIndentStack.pop_back();
        }
        if (mIndentStack.empty() || mIndentStack.back() != indentCols)
        {
          return Fail(ErrorCode::InconsistentIndent, nullptr, "Dedent does not match any previous indent level");
        }
      }
      return ErrorCode::OK;
    }

    // Skip blank / comment lines and position at the first token of the next non-empty line, applying indent stack
    // transitions. Returns false if we reach EOF.
    ErrorCode NextContentLine(bool& hasLine)
    {
      hasLine = false;
      while (!EndOfFile())
      {
        // If we are mid-line, consume to the end so we start clean on the next iteration
        if (mCol != 1)
        {
          SkipToEOL();
          // If we hit EOF while skipping, surface as no more lines
          if (EndOfFile())
          {
            return ErrorCode::OK;
          }
          continue;
        }

        int indentCols = 0;
        bool hasContent = false;
        ErrorCode errorCode = ReadLineIndent(indentCols, hasContent);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }

        if (!hasContent)
        {
          // Blank / comment line, continue to next line
          continue;
        }

        // Apply CoffeeScript indent rules for this new line
        errorCode = ApplyIndentLevel(indentCols);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }

        hasLine = true;
        return ErrorCode::OK;
      }
      return ErrorCode::OK; // EOF, hasLine = false
    }

    static int HexValue(char c)
    {
      if (c >= '0' && c <= '9')
      {
        return c - '0';
      }
      if (c >= 'a' && c <= 'f')
      {
        return 10 + (c - 'a');
      }
      if (c >= 'A' && c <= 'F')
      {
        return 10 + (c - 'A');
      }
      return -1;
    }

    static void AppendUTF8(std::string& out, std::uint32_t codePoint)
    {
      if (codePoint <= 0x7F)
      {
        out.push_back(static_cast<char>(codePoint));
      }
      else if (codePoint <= 0x7FF)
      {
        out.push_back(static_cast<char>(0xC0 | (codePoint >> 6)));
        out.push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
      }
      else if (codePoint <= 0xFFFF)
      {
        out.push_back(static_cast<char>(0xE0 | (codePoint >> 12)));
        out.push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
      }
      else
      {
        out.push_back(static_cast<char>(0xF0 | (codePoint >> 18)));
        out.push_back(static_cast<char>(0x80 | ((codePoint >> 12) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
      }
    }

    bool ValidateUTF8(std::string_view stringView, bool allowLeadingBOM, std::size_t& badIndex, std::size_t& badLine, std::size_t& badCol)
    {
      std::size_t index = 0;
      badLine = 1;
      badCol = 1;

      auto bumpCol = [&](std::uint32_t codePoint) {
        if (codePoint == '\n')
        {
          ++badLine;
          badCol = 1;
        }
        else
        {
          ++badCol;
        }
      };

      while (index < stringView.size())
      {
        unsigned char c = static_cast<unsigned char>(stringView[index]);

        if (
          index == 0 && allowLeadingBOM && stringView.size() >= 3 && c == 0xEF && static_cast<unsigned char>(stringView[1]) == 0xBB &&
          static_cast<unsigned char>(stringView[2]) == 0xBF)
        {
          index += 3;
          continue;
        }
        else if (
          index > 0 && c == 0xEF && index + 2 < stringView.size() && static_cast<unsigned char>(stringView[index + 1]) == 0xBB &&
          static_cast<unsigned char>(stringView[index + 2]) == 0xBF)
        {
          badIndex = index;
          return false; // BOM not at start
        }

        if (c <= 0x7F)
        {
          ++index;
          bumpCol(c);
          continue;
        }

        std::uint32_t codePoint = 0;
        std::size_t length = 0;
        if ((c & 0xE0) == 0xC0)
        {
          length = 2;
          codePoint = c & 0x1F;
        }
        else if ((c & 0xF0) == 0xE0)
        {
          length = 3;
          codePoint = c & 0x0F;
        }
        else if ((c & 0xF8) == 0xF0)
        {
          length = 4;
          codePoint = c & 0x07;
        }
        else
        {
          badIndex = index;
          return false;
        }

        if (index + length > stringView.size())
        {
          badIndex = index;
          return false;
        }
        for (std::size_t k = 1; k < length; ++k)
        {
          unsigned char cc = static_cast<unsigned char>(stringView[index + k]);
          if ((cc & 0xC0) != 0x80)
          {
            badIndex = index;
            return false;
          }
          codePoint = (codePoint << 6) | (cc & 0x3F);
        }

        if (
          (length == 2 && codePoint < 0x80) || (length == 3 && codePoint < 0x800) ||
          (length == 4 && (codePoint < 0x10000 || codePoint > 0x10FFFF)) || (codePoint >= 0xD800 && codePoint <= 0xDFFF))
        {
          badIndex = index;
          return false;
        }

        index += length;
        bumpCol(codePoint);
      }
      return true;
    }

    ErrorCode ParseValue(Value& out, int currentIndent)
    {
      SkipWhitespaceAndComments();

      const auto begin = mPos;
      const auto result = ParseValueBody(out, currentIndent);

      if (result == ErrorCode::OK)
      {
        RecordSource(out, begin);
      }

      return result;
    }

    ErrorCode ParseValueBody(Value& out, int currentIndent)
    {
      char c = Peek();
      if (c == '{')
      {
        return ParseInlineObject(out, currentIndent);
      }
      if (c == '[')
      {
        return ParseArray(out, currentIndent);
      }
      if (c == '"' || c == '\'')
      {
        return ParseQuotedStringOrIndentedObject(out, currentIndent);
      }
      if (IsIdentifierStart(c))
      {
        return parseIdentifierOrIndentedObject(out, currentIndent);
      }
      if (IsNumberStart(c))
      {
        return ParseNumber(out);
      }
      if (EndOfFile())
      {
        return Fail(ErrorCode::UnexpectedEnd, nullptr, "Unexpected end of input while parsing value");
      }
      return Fail(ErrorCode::UnexpectedChar, nullptr, "Unexpected character while parsing value");
    }

    ErrorCode ParseQuotedStringOrIndentedObject(Value& out, int currentIndent)
    {
      const std::size_t savedPos = mPos;
      const std::size_t savedLine = mLine;
      const std::size_t savedColumn = mCol;
      Value quotedValue;
      ErrorCode errorCode = Peek() == '"' ? ParseStringOrTriple(quotedValue) : ParseStringSingle(quotedValue);
      if (errorCode != ErrorCode::OK)
      {
        return errorCode;
      }
      SkipInlineSpaces();
      if (Peek() != ':')
      {
        out = std::move(quotedValue);
        return ErrorCode::OK;
      }
      mPos = savedPos;
      mLine = savedLine;
      mCol = savedColumn;
      Object object;
      errorCode = ParseIndentedObjectBody(object, currentIndent);
      if (errorCode == ErrorCode::OK)
      {
        out = std::move(object);
      }
      return errorCode;
    }

    static bool IsIdentifierStart(char c)
    {
      return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_';
    }

    static bool IsIdentifierChar(char c)
    {
      return IsIdentifierStart(c) || (c >= '0' && c <= '9') || c == '-' || c == '_';
    }

    static bool IsNumberStart(char c)
    {
      return (c >= '0' && c <= '9') || c == '-' || c == '+';
    }

    // Parse identifier or keywords true / false / null or an indent-style object
    ErrorCode parseIdentifierOrIndentedObject(Value& out, int currentIndent)
    {
      const auto savedPos = mPos;
      const auto savedColumn = mCol;
      std::string ident;
      while (IsIdentifierChar(Peek()))
      {
        ident.push_back(Get());
      }
      SkipInlineSpaces();
      if (Peek() == ':')
      {
        // We are at start of an object: key ':' ... possibly multiple pairs
        // Rewind to start of identifier and parse an indented object body.
        // We treat this as an object even at top-level.
        // Reset position so that parseObjectBody can re-read the key.
        mPos = savedPos;
        mCol = savedColumn;
        Object object;
        ErrorCode errorCode = ParseIndentedObjectBody(object, currentIndent);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        out = std::move(object);
        return ErrorCode::OK;
      }

      // Otherwise interpret as an unquoted identifier value
      if (ident == "true")
      {
        out = true;
        return ErrorCode::OK;
      }
      if (ident == "false")
      {
        out = false;
        return ErrorCode::OK;
      }
      if (ident == "null")
      {
        out = nullptr;
        return ErrorCode::OK;
      }
      out = ident; // Unquoted string
      return ErrorCode::OK;
    }

    ErrorCode ParseStringOrTriple(Value& out)
    {
      // We know first char is '"'
      // Look ahead for """
      if (Peek() != '"')
      {
        return Fail(ErrorCode::InternalError, nullptr);
      }
      // Check if next two characters are also '"'
      if (mPos + 2 < mSrc.size() && mSrc[mPos] == '"' && mSrc[mPos + 1] == '"' && mSrc[mPos + 2] == '"')
      {
        return ParseTripleString(out);
      }
      return ParseStringDouble(out);
    }

    ErrorCode ParseStringDouble(Value& out)
    {
      if (!Match('"'))
      {
        return Fail(ErrorCode::InternalError, nullptr);
      }
      std::string result;
      while (!EndOfFile())
      {
        char c = Get();
        if (c == '"')
        {
          out = std::move(result);
          return ErrorCode::OK;
        }
        if (c == '\\')
        {
          if (EndOfFile())
          {
            return Fail(ErrorCode::UnterminatedString, nullptr);
          }
          char e = Get();
          switch (e)
          {
            case '"': result.push_back('"'); break;
            case '\\': result.push_back('\\'); break;
            case 'n': result.push_back('\n'); break;
            case 'r': result.push_back('\r'); break;
            case 't': result.push_back('\t'); break;
            case 'u':
            {
              // Unicode escape \uXXXX with optional surrogate pair
              std::uint32_t codePoint = 0;
              for (std::size_t index = 0; index < 4; ++index)
              {
                if (EndOfFile())
                {
                  return Fail(ErrorCode::InvalidEscape, nullptr, "Incomplete unicode escape");
                }
                int hv = HexValue(Get());
                if (hv < 0)
                {
                  return Fail(ErrorCode::InvalidEscape, nullptr, "Invalid hex in unicode escape");
                }
                codePoint = (codePoint << 4) | static_cast<std::uint32_t>(hv);
              }
              if (codePoint >= 0xD800 && codePoint <= 0xDBFF)
              {
                // Expect low surrogate
                if (!(Peek() == '\\' && mPos + 1 < mSrc.size() && mSrc[mPos + 1] == 'u'))
                {
                  return Fail(ErrorCode::InvalidEscape, nullptr, "Unpaired surrogate");
                }
                Get(); // '\'
                Get(); // 'u'
                std::uint32_t lowSurrogate = 0;
                for (std::size_t index = 0; index < 4; ++index)
                {
                  if (EndOfFile())
                    return Fail(ErrorCode::InvalidEscape, nullptr, "Incomplete unicode escape");
                  int hexValue = HexValue(Get());
                  if (hexValue < 0)
                  {
                    return Fail(ErrorCode::InvalidEscape, nullptr, "Invalid hex in unicode escape");
                  }
                  lowSurrogate = (lowSurrogate << 4) | static_cast<std::uint32_t>(hexValue);
                }
                if (lowSurrogate < 0xDC00 || lowSurrogate > 0xDFFF)
                {
                  return Fail(ErrorCode::InvalidEscape, nullptr, "Invalid low surrogate");
                }
                codePoint = 0x10000 + (((codePoint - 0xD800) << 10) | (lowSurrogate - 0xDC00));
              }
              else if (codePoint >= 0xDC00 && codePoint <= 0xDFFF)
              {
                return Fail(ErrorCode::InvalidEscape, nullptr, "Unpaired surrogate");
              }
              AppendUTF8(result, codePoint);
              break;
            }
            default: return Fail(ErrorCode::InvalidEscape, nullptr, "Invalid escape in string");
          }
        }
        else if (c == '\n' || c == '\r')
        {
          return Fail(ErrorCode::UnterminatedString, nullptr, "Newline in string literal");
        }
        else
        {
          result.push_back(c);
        }
      }
      return Fail(ErrorCode::UnterminatedString, nullptr, "Unterminated string literal");
    }

    ErrorCode ParseStringSingle(Value& out)
    {
      if (!Match('\''))
      {
        return Fail(ErrorCode::InternalError, nullptr);
      }
      std::string result;
      while (!EndOfFile())
      {
        char c = Get();
        if (c == '\'')
        {
          out = std::move(result);
          return ErrorCode::OK;
        }
        if (c == '\\')
        {
          if (EndOfFile())
          {
            return Fail(ErrorCode::UnterminatedString, nullptr);
          }
          char e = Get();
          switch (e)
          {
            case '\'': result.push_back('\''); break;
            case '\\': result.push_back('\\'); break;
            case 'n': result.push_back('\n'); break;
            case 'r': result.push_back('\r'); break;
            case 't': result.push_back('\t'); break;
            case 'u':
            {
              std::uint32_t codePoint = 0;
              for (std::size_t index = 0; index < 4; ++index)
              {
                if (EndOfFile())
                {
                  return Fail(ErrorCode::InvalidEscape, nullptr, "Incomplete unicode escape");
                }
                int hexValue = HexValue(Get());
                if (hexValue < 0)
                {
                  return Fail(ErrorCode::InvalidEscape, nullptr, "Invalid hex in unicode escape");
                }
                codePoint = (codePoint << 4) | static_cast<std::uint32_t>(hexValue);
              }
              if (codePoint >= 0xD800 && codePoint <= 0xDBFF)
              {
                if (!(Peek() == '\\' && mPos + 1 < mSrc.size() && mSrc[mPos + 1] == 'u'))
                {
                  return Fail(ErrorCode::InvalidEscape, nullptr, "Unpaired surrogate");
                }
                Get();
                Get();
                std::uint32_t lowSurrogate = 0;
                for (std::size_t index = 0; index < 4; ++index)
                {
                  if (EndOfFile())
                  {
                    return Fail(ErrorCode::InvalidEscape, nullptr, "Incomplete unicode escape");
                  }
                  int hexValue = HexValue(Get());
                  if (hexValue < 0)
                  {
                    return Fail(ErrorCode::InvalidEscape, nullptr, "Invalid hex in unicode escape");
                  }
                  lowSurrogate = (lowSurrogate << 4) | static_cast<std::uint32_t>(hexValue);
                }
                if (lowSurrogate < 0xDC00 || lowSurrogate > 0xDFFF)
                {
                  return Fail(ErrorCode::InvalidEscape, nullptr, "Invalid low surrogate");
                }
                codePoint = 0x10000 + (((codePoint - 0xD800) << 10) | (lowSurrogate - 0xDC00));
              }
              else if (codePoint >= 0xDC00 && codePoint <= 0xDFFF)
              {
                return Fail(ErrorCode::InvalidEscape, nullptr, "Unpaired surrogate");
              }
              AppendUTF8(result, codePoint);
              break;
            }
            default: return Fail(ErrorCode::InvalidEscape, nullptr, "Invalid escape in string");
          }
        }
        else if (c == '\n' || c == '\r')
        {
          return Fail(ErrorCode::UnterminatedString, nullptr, "Newline in string literal");
        }
        else
        {
          result.push_back(c);
        }
      }
      return Fail(ErrorCode::UnterminatedString, nullptr, "Unterminated string literal");
    }

    ErrorCode ParseTripleString(Value& out)
    {
      // Consume initial """
      if (!(Match('"') && Match('"') && Match('"')))
      {
        return Fail(ErrorCode::InternalError, nullptr);
      }
      std::string result;
      while (!EndOfFile())
      {
        if (Peek() == '"' && mPos + 2 < mSrc.size() && mSrc[mPos + 1] == '"' && mSrc[mPos + 2] == '"')
        {
          // End
          Get();
          Get();
          Get();
          out = std::move(result);
          return ErrorCode::OK;
        }
        result.push_back(Get());
      }
      return Fail(ErrorCode::UnterminatedTripleString, nullptr, "Unterminated triple string literal");
    }

    ErrorCode ParseNumber(Value& out)
    {
      std::size_t start = mPos;
      bool hasDot = false;
      bool hasExp = false;

      if (Peek() == '+' || Peek() == '-')
      {
        Get();
      }

      while (!EndOfFile())
      {
        char c = Peek();
        if (c >= '0' && c <= '9')
        {
          Get();
        }
        else if (c == '.' && !hasDot)
        {
          hasDot = true;
          Get();
        }
        else if ((c == 'e' || c == 'E') && !hasExp)
        {
          hasExp = true;
          Get();
          if (Peek() == '+' || Peek() == '-')
          {
            Get();
          }
        }
        else
        {
          break;
        }
      }

      std::string_view stringView(mSrc.data() + start, mPos - start);
      const char* first = stringView.data();
      const char* last = first + stringView.size();
      if (first != last && *first == '+')
      {
        ++first;
      }
      double value = 0.0;
      const auto result = std::from_chars(first, last, value, std::chars_format::general);
      if (first == last || result.ec != std::errc{} || result.ptr != last || !std::isfinite(value))
      {
        return Fail(ErrorCode::InvalidNumber, nullptr, "Invalid number literal");
      }
      out = value;
      return ErrorCode::OK;
    }

    ErrorCode ParseInlineObject(Value& out, int currentIndent)
    {
      if (!CanEnterContainer())
      {
        return ErrorCode::ResourceLimit;
      }
      const ContainerDepthGuard depthGuard(mContainerDepth);
      if (!Match('{'))
      {
        return Fail(ErrorCode::InternalError, nullptr);
      }
      Object object;
      SkipWhitespaceAndComments();
      if (Match('}'))
      {
        out = std::move(object);
        return ErrorCode::OK;
      }
      while (true)
      {
        const LocationEntry keyLocation = Location();
        std::string key;
        SourceSpan keySource;
        ErrorCode errorCode = ParseKey(key, &keySource);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        errorCode = CheckDuplicateKey(object, key, keyLocation);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        SkipWhitespaceAndComments();
        if (!Match(':'))
        {
          return Fail(ErrorCode::UnexpectedChar, nullptr, "Expected ':' in object");
        }
        SkipWhitespaceAndComments();
        Value value;
        errorCode = ParseValue(value, currentIndent);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        RecordKeySource(value, keySource);
        object.emplace(std::move(key), std::move(value));
        SkipWhitespaceAndComments();
        if (Match('}'))
        {
          break;
        }
        if (Peek() == '#')
        {
          SkipToEOL();
          SkipWhitespaceAndComments();
          if (Match('}'))
          {
            break;
          }
        }
        if (Match(','))
        {
          SkipWhitespaceAndComments();
          continue;
        }
        return Fail(ErrorCode::UnexpectedChar, nullptr, "Expected ',' or '}' in object");
      }
      out = std::move(object);
      return ErrorCode::OK;
    }

    ErrorCode ParseIndentedObjectBody(Object& object, int parentIndent)
    {
      if (!CanEnterContainer())
      {
        return ErrorCode::ResourceLimit;
      }
      const ContainerDepthGuard depthGuard(mContainerDepth);
      // We assume we're currently on the line that already has the first key at indent == mIndentStack.back() (>=
      // parentIndent). The object spans lines at the current indent; deeper indents belong to child values.
      int bodyIndent = -1; // Will be set after parsing first key

      while (true)
      {
        // Set bodyIndent on first iteration
        if (bodyIndent == -1)
        {
          bodyIndent = mIndentStack.back();
        }
        SkipInlineSpaces();
        if (EndOfFile())
        {
          break;
        }

        // Skip blank / comment lines inside an object body
        if (Peek() == '#' || Peek() == '\r' || Peek() == '\n')
        {
          bool hasLine = false;
          ErrorCode ec = NextContentLine(hasLine);
          if (ec != ErrorCode::OK)
          {
            return ec;
          }
          if (!hasLine)
          {
            break; // EOF
          }
          if (mIndentStack.back() < bodyIndent)
          {
            break;
          }
          continue;
        }

        char c = Peek();
        if (!IsIdentifierStart(c) && c != '"' && c != '\'')
        {
          // Probably end of this block (dedent handled by caller)
          break;
        }

        const LocationEntry keyLocation = Location();
        std::string key;
        SourceSpan keySource;
        ErrorCode errorCode = ParseKey(key, &keySource);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        errorCode = CheckDuplicateKey(object, key, keyLocation);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }

        SkipInlineSpaces();
        if (!Match(':'))
        {
          return Fail(ErrorCode::UnexpectedChar, nullptr, "Expected ':' in object pair");
        }
        SkipInlineSpaces();

        // Comment immediately after ':' -> treat as block value starting on next line
        if (Peek() == '#')
        {
          SkipToEOL();

          bool hasLine = false;
          ErrorCode errorCode2 = NextContentLine(hasLine);
          if (errorCode2 != ErrorCode::OK)
          {
            return errorCode2;
          }

          if (!hasLine)
          {
            return Fail(ErrorCode::InconsistentIndent, nullptr, "Expected indented block after ':'");
          }

          if (mIndentStack.back() <= bodyIndent)
          {
            return Fail(ErrorCode::InconsistentIndent, nullptr, "Expected deeper indentation for block value");
          }

          Value value;
          errorCode2 = ParseValue(value, mIndentStack.back());
          if (errorCode2 != ErrorCode::OK)
          {
            return errorCode2;
          }
          RecordKeySource(value, keySource);
          object.emplace(std::move(key), std::move(value));

          if (mIndentStack.back() < bodyIndent)
          {
            break;
          }
          continue;
        }

        Value value;

        // If nothing else on the line -> block value (object / array) on next indented line
        if (Peek() == '\r' || Peek() == '\n')
        {
          // Consume EOL
          char c2 = Get();
          if (c2 == '\r' && Peek() == '\n')
          {
            Get();
          }

          bool hasLine = false;
          ErrorCode errorCode2 = NextContentLine(hasLine);
          if (errorCode2 != ErrorCode::OK)
          {
            return errorCode2;
          }

          if (!hasLine)
          {
            return Fail(ErrorCode::InconsistentIndent, nullptr, "Expected indented block after ':'");
          }

          // Indent stack already updated; current indent = mIndentStack.back()
          if (mIndentStack.back() <= bodyIndent)
          {
            return Fail(ErrorCode::InconsistentIndent, nullptr, "Expected deeper indentation for block value");
          }

          // Recursively parse value at new indent level
          errorCode2 = ParseValue(value, mIndentStack.back());
          if (errorCode2 != ErrorCode::OK)
          {
            return errorCode2;
          }

          // After parsing block value, check if we've dedented (nextContentLine was called inside the recursive parse)
          RecordKeySource(value, keySource);
          object.emplace(std::move(key), std::move(value));

          // Check current indent level - if we've dedented out of this object, we're done
          if (mIndentStack.back() < bodyIndent)
          {
            break;
          }

          // Otherwise continue parsing more keys at this level
          continue;
        }
        else
        {
          // Inline value on same line
          errorCode = ParseValue(value, parentIndent);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
        }

        RecordKeySource(value, keySource);
        object.emplace(std::move(key), std::move(value));

        // End of line or another entry on same line (comma separated)
        SkipInlineSpaces();
        if (Peek() == '#')
        {
          SkipToEOL();
          bool hasLine = false;
          ErrorCode errorCode2 = NextContentLine(hasLine);
          if (errorCode2 != ErrorCode::OK)
          {
            return errorCode2;
          }
          if (!hasLine)
          {
            break; // EOF
          }
          if (mIndentStack.back() < bodyIndent)
          {
            break;
          }
          continue;
        }
        if (Match(','))
        {
          SkipInlineSpaces();
          // Next key / value on same line
          continue;
        }

        // If we hit newline, move to next content line and see whether indent still belongs to this object or a parent
        if (Peek() == '\r' || Peek() == '\n')
        {
          char c2 = Get();
          if (c2 == '\r' && Peek() == '\n')
          {
            Get();
          }

          bool hasLine = false;
          ErrorCode errorCode2 = NextContentLine(hasLine);
          if (errorCode2 != ErrorCode::OK)
          {
            return errorCode2;
          }
          if (!hasLine)
          {
            break; // EOF
          }

          // If we dedented to or below the body's indent, the object ends
          if (mIndentStack.back() < bodyIndent)
          {
            break;
          }

          // Still inside this object; loop continues with new line
          continue;
        }

        // Otherwise (e.g., end of file or delimiters), let caller decide
        break;
      }

      return ErrorCode::OK;
    }

    ErrorCode ParseArray(Value& out, int parentIndent)
    {
      if (!CanEnterContainer())
      {
        return ErrorCode::ResourceLimit;
      }
      const ContainerDepthGuard depthGuard(mContainerDepth);
      if (!Match('['))
      {
        return Fail(ErrorCode::InternalError, nullptr);
      }

      Array array;
      std::optional<std::size_t> closingEnd;
      const auto closeArray = [&]()
      {
        if (!Match(']'))
        {
          return false;
        }
        closingEnd = mPos;
        return true;
      };

      // Check if this is a multiline array (newline after '[')
      SkipInlineSpaces();
      bool isMultiline = (Peek() == '\r' || Peek() == '\n');

      if (isMultiline)
      {
        // Consume the newline and position at first element
        char c = Get();
        if (c == '\r' && Peek() == '\n')
        {
          Get();
        }

        bool hasLine = false;
        ErrorCode errorCode = NextContentLine(hasLine);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        if (!hasLine || Peek() == ']')
        {
          // Empty array or just closing bracket
          if (Peek() == ']')
          {
            (void)closeArray();
          }
          out = std::move(array);
          mArrayClosingEnd = closingEnd;
          return ErrorCode::OK;
        }

        // Parse multiline array elements
        int arrayIndent = mIndentStack.back();
        while (true)
        {
          if (Peek() == ']')
          {
            (void)closeArray();
            break;
          }

          Value value;
          errorCode = ParseValue(value, arrayIndent);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          array.push_back(std::move(value));

          // After parsing, check if we've dedented out of the array
          if (mIndentStack.back() < arrayIndent)
          {
            break;
          }

          SkipInlineSpaces();
          if (Peek() == '#')
          {
            SkipToEOL();
            bool hasLine2 = false;
            errorCode = NextContentLine(hasLine2);
            if (errorCode != ErrorCode::OK)
            {
              return errorCode;
            }
            if (!hasLine2)
            {
              break;
            }
            if (mIndentStack.back() < arrayIndent)
            {
              break;
            }
            continue;
          }
          if (closeArray())
          {
            break;
          }

          // Check for comma
          if (Match(','))
          {
            SkipWhitespaceAndComments();
            continue;
          }

          // Check for newline
          if (Peek() == '\r' || Peek() == '\n')
          {
            c = Get();
            if (c == '\r' && Peek() == '\n')
            {
              Get();
            }

            // Use nextContentLine to properly handle indent stack
            hasLine = false;
            errorCode = NextContentLine(hasLine);
            if (errorCode != ErrorCode::OK)
            {
              return errorCode;
            }
            if (!hasLine)
            {
              break;
            }

            // Check if we've dedented out of the array
            if (mIndentStack.back() < arrayIndent)
            {
              break;
            }

            // Continue to parse next element
            continue;
          }

          return Fail(ErrorCode::UnexpectedChar, nullptr, "Expected ',' or ']' or newline in multiline array");
        }

        if (Peek() == ']')
        {
          (void)closeArray();
        }
      }
      else
      {
        // Inline array
        SkipWhitespaceAndComments();
        if (closeArray())
        {
          out = std::move(array);
          mArrayClosingEnd = closingEnd;
          return ErrorCode::OK;
        }

        while (true)
        {
          if (Peek() == ']')
          {
            (void)closeArray();
            break;
          }

          Value value;
          ErrorCode errorCode = ParseValue(value, parentIndent);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          array.push_back(std::move(value));

          SkipWhitespaceAndComments();
          if (closeArray())
          {
            break;
          }
          if (!Match(','))
          {
            return Fail(ErrorCode::UnexpectedChar, nullptr, "Expected ',' or ']' in inline array");
          }
          SkipWhitespaceAndComments();
        }
      }

      out = std::move(array);
      mArrayClosingEnd = closingEnd;
      return ErrorCode::OK;
    }

    ErrorCode ParseKey(std::string& outKey, SourceSpan* source = nullptr)
    {
      SkipInlineSpaces();
      const auto begin = mPos;
      const auto result = ParseKeyBody(outKey);
      if (result == ErrorCode::OK && source && mOptions.trackSourceLocations)
      {
        *source = {SourcePositionAt(begin), SourcePositionAt(mPos)};
      }
      return result;
    }

    ErrorCode ParseKeyBody(std::string& outKey)
    {
      char c = Peek();
      if (c == '"')
      {
        Value tempValue;
        ErrorCode errorCode = ParseStringDouble(tempValue);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        outKey = std::get<std::string>(tempValue);
        return ErrorCode::OK;
      }
      if (c == '\'')
      {
        Value tempValue;
        ErrorCode errorCode = ParseStringSingle(tempValue);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        outKey = std::get<std::string>(tempValue);
        return ErrorCode::OK;
      }
      if (!IsIdentifierStart(c))
      {
        return Fail(ErrorCode::UnexpectedChar, nullptr, "Expected identifier or string as key");
      }
      std::string ident;
      while (IsIdentifierChar(Peek()))
      {
        ident.push_back(Get());
      }
      outKey = std::move(ident);
      return ErrorCode::OK;
    }
  };

  // Lossless edits use owned key/index paths. An empty path denotes the root.
  // Source locations continue to refer to the original input after edits.
  enum class RemovedCommentPolicy : std::uint8_t { Preserve, Discard };
  enum class CommentIndentation : std::uint8_t { MatchDestination, PreserveOriginal };

  struct LosslessEditOptions
  {
    RemovedCommentPolicy removedComments = RemovedCommentPolicy::Preserve;
    CommentIndentation commentIndentation = CommentIndentation::MatchDestination;
    int indentWidth = 2; // Match the WriteOptions used when saving
  };

  namespace detail
  {
    inline bool LosslessError(Error* error, ErrorCode code, const ValuePath& path,
                              std::string_view message, const Value* value = nullptr)
    {
      if (error)
      {
        *error = {};
        error->code = code;
        error->message = FormatValuePath(path) + ": " + std::string(message);
        if (value && value->Source())
        {
          const auto& source = *value->Source();
          error->where = {source.valueSpan.begin.line, source.valueSpan.begin.column};
          if (source.filename)
          {
            error->filename = *source.filename;
          }
        }
      }
      return false;
    }

    inline bool CheckLosslessTree(const LosslessValue& node, ValuePath& path,
                                   Error* error, bool compareSemantic)
    {
      if (node.value.isObject())
      {
        if (!node.arrayItems.empty())
        {
          return LosslessError(error, ErrorCode::InvalidLosslessTree, path,
                               "Object contains ordered array children", &node.value);
        }
        Object seen;
        for (const auto& [key, child] : node.objectItems)
        {
          path.emplace_back(key);
          if (!seen.emplace(key, nullptr).second)
          {
            return LosslessError(error, ErrorCode::DuplicateKey, path,
                                 "Duplicate ordered object member", &child.value);
          }
          if (!CheckLosslessTree(child, path, error, compareSemantic))
          {
            return false;
          }
          if (compareSemantic)
          {
            const auto found = node.value.asObject().find(key);
            if (found == node.value.asObject().end() || found->second != child.value)
            {
              return LosslessError(error, ErrorCode::InvalidLosslessTree, path,
                                   "Semantic and ordered object values differ", &child.value);
            }
          }
          path.pop_back();
        }
        if ((compareSemantic && node.value.asObject().size() != node.objectItems.size()) ||
            (!compareSemantic && node.objectItems.empty() && !node.value.asObject().empty()))
        {
          return LosslessError(error, ErrorCode::InvalidLosslessTree, path,
                               "Semantic members lack ordered children; use MakeLossless for semantic-only values", &node.value);
        }
      }
      else if (node.value.isArray())
      {
        if (!node.objectItems.empty())
        {
          return LosslessError(error, ErrorCode::InvalidLosslessTree, path,
                               "Array contains ordered object members", &node.value);
        }
        if ((compareSemantic && node.value.asArray().size() != node.arrayItems.size()) ||
            (!compareSemantic && node.arrayItems.empty() && !node.value.asArray().empty()))
        {
          return LosslessError(error, ErrorCode::InvalidLosslessTree, path,
                               "Semantic array lacks matching ordered children; use MakeLossless for semantic-only values", &node.value);
        }
        for (std::size_t index = 0; index < node.arrayItems.size(); ++index)
        {
          path.emplace_back(index);
          if (!CheckLosslessTree(node.arrayItems[index], path, error, compareSemantic))
          {
            return false;
          }
          if (compareSemantic && node.value.asArray()[index] != node.arrayItems[index].value)
          {
            return LosslessError(error, ErrorCode::InvalidLosslessTree, path,
                                 "Semantic and ordered array values differ", &node.arrayItems[index].value);
          }
          path.pop_back();
        }
      }
      else if (!node.arrayItems.empty() || !node.objectItems.empty())
      {
        return LosslessError(error, ErrorCode::InvalidLosslessTree, path,
                             "Scalar contains ordered children", &node.value);
      }
      return true;
    }

    inline void SynchronizeLosslessTree(LosslessValue& node)
    {
      if (node.value.isObject())
      {
        Object members;
        members.reserve(node.objectItems.size());
        for (auto& [key, child] : node.objectItems)
        {
          SynchronizeLosslessTree(child);
          members.emplace(key, child.value);
        }
        node.value.asObject() = std::move(members); // Retain the container's original source locations
      }
      else if (node.value.isArray())
      {
        Array items;
        items.reserve(node.arrayItems.size());
        for (auto& child : node.arrayItems)
        {
          SynchronizeLosslessTree(child);
          items.push_back(child.value);
        }
        node.value.asArray() = std::move(items);
      }
    }

    inline const LosslessValue* FindLosslessPath(const LosslessValue& root,
                                                const ValuePath& path, Error* error)
    {
      const LosslessValue* current = &root;
      ValuePath traversed;
      for (const auto& segment : path)
      {
        traversed.push_back(segment);
        if (const auto* key = std::get_if<std::string>(&segment))
        {
          if (!current->value.isObject())
          {
            LosslessError(error, ErrorCode::TypeMismatch, traversed, "Expected an object", &current->value);
            return nullptr;
          }
          const auto found = std::find_if(current->objectItems.begin(), current->objectItems.end(),
            [&](const auto& member) { return member.first == *key; });
          if (found == current->objectItems.end())
          {
            LosslessError(error, ErrorCode::MissingMember, traversed, "Object member does not exist", &current->value);
            return nullptr;
          }
          current = &found->second;
        }
        else
        {
          if (!current->value.isArray())
          {
            LosslessError(error, ErrorCode::TypeMismatch, traversed, "Expected an array", &current->value);
            return nullptr;
          }
          const auto index = std::get<std::size_t>(segment);
          if (index >= current->arrayItems.size())
          {
            LosslessError(error, ErrorCode::OutOfRange, traversed, "Array index is out of range", &current->value);
            return nullptr;
          }
          current = &current->arrayItems[index];
        }
      }
      return current;
    }
  }

  inline bool ValidateLosslessTree(const LosslessValue& node, Error* error = nullptr)
  {
    if (error)
    {
      *error = {};
    }
    ValuePath path;
    return detail::CheckLosslessTree(node, path, error, true);
  }

  // Explicitly choose ordered children as authoritative. Contradictory child
  // kinds, duplicate keys, and nonempty semantic-only containers are rejected.
  inline bool RebuildLosslessTree(LosslessValue& node, Error* error = nullptr)
  {
    if (error)
    {
      *error = {};
    }
    ValuePath path;
    if (!detail::CheckLosslessTree(node, path, error, false))
    {
      return false;
    }
    LosslessValue rebuilt = node;
    detail::SynchronizeLosslessTree(rebuilt);
    node = std::move(rebuilt);
    return true;
  }

  // Semantic objects have no order to preserve. Sort by key by default for
  // reproducible output; false uses the unordered_map's iteration order.
  inline LosslessValue MakeLossless(const Value& value, bool sortObjectKeys = true)
  {
    LosslessValue result;
    result.value = value;
    if (value.isArray())
    {
      for (const auto& child : value.asArray())
      {
        result.arrayItems.push_back(MakeLossless(child, sortObjectKeys));
      }
    }
    else if (value.isObject())
    {
      for (const auto& [key, child] : value.asObject())
      {
        result.objectItems.emplace_back(key, MakeLossless(child, sortObjectKeys));
      }
      if (sortObjectKeys)
      {
        std::sort(result.objectItems.begin(), result.objectItems.end(),
          [](const auto& left, const auto& right) { return left.first < right.first; });
      }
    }
    return result;
  }

  class LosslessDocumentEditor
  {
  public:
    static constexpr std::size_t Append = std::numeric_limits<std::size_t>::max();

    explicit LosslessDocumentEditor(LosslessValue& root, LosslessEditOptions options = {})
      : mRoot(root), mOptions(options) {}

    const LosslessValue& Root() const noexcept { return mRoot; }

    const LosslessValue* Find(const ValuePath& path, Error* error = nullptr) const
    {
      if (!ValidateLosslessTree(mRoot, error))
      {
        return nullptr;
      }
      return detail::FindLosslessPath(mRoot, path, error);
    }

    bool InsertMember(const ValuePath& path, std::string key, LosslessValue value,
                      std::size_t position = Append, Error* error = nullptr)
    {
      if (!ValidateLosslessTree(value, error))
      {
        return false;
      }
      return Edit(path, error, [&](LosslessValue& object)
      {
        if (!object.value.isObject())
        {
          return Fail(error, ErrorCode::TypeMismatch, path, "Expected an object");
        }
        if (MemberIndex(object, key) != Append)
        {
          return Fail(error, ErrorCode::DuplicateKey, path, "Object member already exists");
        }
        if (position == Append)
        {
          position = object.objectItems.size();
        }
        if (position > object.objectItems.size())
        {
          return Fail(error, ErrorCode::OutOfRange, path, "Member insertion index is out of range");
        }
        object.objectItems.insert(object.objectItems.begin() + static_cast<std::ptrdiff_t>(position),
                                  {std::move(key), std::move(value)});
        return true;
      });
    }

    // Original comments survive replacement, including comments in removed
    // descendants. Incoming comments are retained as well. Fresh replacement
    // values have no source locations. Copies retain their original source locations.
    bool Replace(const ValuePath& path, LosslessValue replacement, Error* error = nullptr)
    {
      if (!ValidateLosslessTree(replacement, error))
      {
        return false;
      }
      const int indent = CommentIndent(path);
      return Edit(path, error, [&](LosslessValue& original)
      {
        PreserveReplacementComments(original, replacement, indent);
        original = std::move(replacement);
        return true;
      });
    }

    bool ReplaceMember(const ValuePath& path, std::string key, LosslessValue replacement, Error* error = nullptr)
    {
      auto child = path;
      child.emplace_back(std::move(key));
      return Replace(child, std::move(replacement), error);
    }

    bool RenameMember(const ValuePath& path, std::string_view oldKey, std::string newKey, Error* error = nullptr)
    {
      return Edit(path, error, [&](LosslessValue& object)
      {
        if (!object.value.isObject())
        {
          return Fail(error, ErrorCode::TypeMismatch, path, "Expected an object");
        }
        const auto index = MemberIndex(object, oldKey);
        if (index == Append)
        {
          return Fail(error, ErrorCode::MissingMember, path, "Object member does not exist");
        }
        if (oldKey != newKey && MemberIndex(object, newKey) != Append)
        {
          return Fail(error, ErrorCode::DuplicateKey, path, "Object member already exists");
        }
        object.objectItems[index].first = std::move(newKey);
        return true;
      });
    }

    bool Remove(const ValuePath& path, Error* error = nullptr)
    {
      if (path.empty())
      {
        return Fail(error, ErrorCode::InvalidPath, path, "Cannot remove the document root");
      }
      auto parentPath = path;
      const auto segment = parentPath.back();
      parentPath.pop_back();
      const int indent = CommentIndent(path);
      return Edit(parentPath, error, [&](LosslessValue& parent)
      {
        if (const auto* key = std::get_if<std::string>(&segment))
        {
          if (!parent.value.isObject())
          {
            return Fail(error, ErrorCode::TypeMismatch, path, "Expected an object");
          }
          const auto index = MemberIndex(parent, *key);
          if (index == Append)
          {
            return Fail(error, ErrorCode::MissingMember, path, "Object member does not exist");
          }
          auto comments = RemovedComments(parent.objectItems[index].second, indent);
          parent.objectItems.erase(parent.objectItems.begin() + static_cast<std::ptrdiff_t>(index));
          if (index < parent.objectItems.size())
          {
            Prepend(parent.objectItems[index].second.leadingComments, std::move(comments));
          }
          else
          {
            RelocateRemovedTailComments(parent, parentPath, std::move(comments));
          }
        }
        else
        {
          if (!parent.value.isArray())
          {
            return Fail(error, ErrorCode::TypeMismatch, path, "Expected an array");
          }
          const auto index = std::get<std::size_t>(segment);
          if (index >= parent.arrayItems.size())
          {
            return Fail(error, ErrorCode::OutOfRange, path, "Array index is out of range");
          }
          auto comments = RemovedComments(parent.arrayItems[index], indent);
          parent.arrayItems.erase(parent.arrayItems.begin() + static_cast<std::ptrdiff_t>(index));
          if (index < parent.arrayItems.size())
          {
            Prepend(parent.arrayItems[index].leadingComments, std::move(comments));
          }
          else
          {
            RelocateRemovedTailComments(parent, parentPath, std::move(comments));
          }
        }
        return true;
      });
    }

    bool RemoveMember(const ValuePath& path, std::string key, Error* error = nullptr)
    {
      auto child = path;
      child.emplace_back(std::move(key));
      return Remove(child, error);
    }

    bool MoveMember(const ValuePath& path, std::string_view key, std::size_t position, Error* error = nullptr)
    {
      return Edit(path, error, [&](LosslessValue& object)
      {
        if (!object.value.isObject())
        {
          return Fail(error, ErrorCode::TypeMismatch, path, "Expected an object");
        }
        const auto index = MemberIndex(object, key);
        if (index == Append)
        {
          return Fail(error, ErrorCode::MissingMember, path, "Object member does not exist");
        }
        if (position == Append)
        {
          position = object.objectItems.size() - 1;
        }
        if (position >= object.objectItems.size())
        {
          return Fail(error, ErrorCode::OutOfRange, path, "Member destination index is out of range");
        }
        MoveItem(object.objectItems, index, position);
        return true;
      });
    }

    bool InsertArrayItem(const ValuePath& path, LosslessValue value,
                         std::size_t position = Append, Error* error = nullptr)
    {
      if (!ValidateLosslessTree(value, error))
      {
        return false;
      }
      return Edit(path, error, [&](LosslessValue& array)
      {
        if (!array.value.isArray())
        {
          return Fail(error, ErrorCode::TypeMismatch, path, "Expected an array");
        }
        if (position == Append)
        {
          position = array.arrayItems.size();
        }
        if (position > array.arrayItems.size())
        {
          return Fail(error, ErrorCode::OutOfRange, path, "Array insertion index is out of range");
        }
        array.arrayItems.insert(array.arrayItems.begin() + static_cast<std::ptrdiff_t>(position), std::move(value));
        return true;
      });
    }

    bool ReplaceArrayItem(const ValuePath& path, std::size_t index, LosslessValue replacement, Error* error = nullptr)
    {
      auto child = path;
      child.emplace_back(index);
      return Replace(child, std::move(replacement), error);
    }

    bool RemoveArrayItem(const ValuePath& path, std::size_t index, Error* error = nullptr)
    {
      auto child = path;
      child.emplace_back(index);
      return Remove(child, error);
    }

    // The destination is the final zero-based index, not a pre-removal gap
    bool MoveArrayItem(const ValuePath& path, std::size_t index, std::size_t position, Error* error = nullptr)
    {
      return Edit(path, error, [&](LosslessValue& array)
      {
        if (!array.value.isArray())
        {
          return Fail(error, ErrorCode::TypeMismatch, path, "Expected an array");
        }
        if (position == Append && !array.arrayItems.empty())
        {
          position = array.arrayItems.size() - 1;
        }
        if (index >= array.arrayItems.size() || position >= array.arrayItems.size())
        {
          return Fail(error, ErrorCode::OutOfRange, path, "Array index is out of range");
        }
        MoveItem(array.arrayItems, index, position);
        return true;
      });
    }

  private:
    LosslessValue& mRoot;
    LosslessEditOptions mOptions;

    static bool Fail(Error* error, ErrorCode code, const ValuePath& path, std::string_view message)
    {
      return detail::LosslessError(error, code, path, message);
    }

    template<class Operation>
    bool Edit(const ValuePath& path, Error* error, Operation&& operation)
    {
      if (!ValidateLosslessTree(mRoot, error))
      {
        return false;
      }
      if (mOptions.indentWidth < 1)
      {
        return Fail(error, ErrorCode::OutOfRange, path, "Comment indentation width must be positive");
      }
      LosslessValue candidate = mRoot;
      auto* node = const_cast<LosslessValue*>(detail::FindLosslessPath(candidate, path, error));
      if (!node || !operation(*node))
      {
        return false;
      }
      detail::SynchronizeLosslessTree(candidate);
      mRoot = std::move(candidate);
      if (error)
      {
        *error = {};
      }
      return true;
    }

    static std::size_t MemberIndex(const LosslessValue& object, std::string_view key)
    {
      for (std::size_t index = 0; index < object.objectItems.size(); ++index)
      {
        if (object.objectItems[index].first == key)
        {
          return index;
        }
      }
      return Append;
    }

    template<class Item>
    static void MoveItem(std::vector<Item>& items, std::size_t from, std::size_t to)
    {
      if (from == to)
      {
        return;
      }
      Item item = std::move(items[from]);
      items.erase(items.begin() + static_cast<std::ptrdiff_t>(from));
      items.insert(items.begin() + static_cast<std::ptrdiff_t>(to), std::move(item));
    }

    static void AppendComments(std::vector<LosslessComment>& target, std::vector<LosslessComment> comments)
    {
      for (auto& comment : comments)
      {
        target.push_back(std::move(comment));
      }
    }

    static void Prepend(std::vector<LosslessComment>& target, std::vector<LosslessComment> comments)
    {
      AppendComments(comments, std::move(target));
      target = std::move(comments);
    }

    static void RelocateRemovedTailComments(LosslessValue& parent, const ValuePath& parentPath,
                                            std::vector<LosslessComment> comments)
    {
      // Keep removed-child trivia before the container's closing delimiter and
      // its existing closing comments. Trailing comments belong outside it.
      // Implicit objects have no closing delimiter unless the writer needs
      // braces for closing trivia, an inline comment, or an array element.
      const bool hasClosingDelimiter = parent.value.isArray() || !parent.closingComments.empty() ||
        !parent.inlineComment.empty() ||
        (!parentPath.empty() && std::holds_alternative<std::size_t>(parentPath.back()));

      if (hasClosingDelimiter)
      {
        Prepend(parent.closingComments, std::move(comments));
      }
      else
      {
        Prepend(parent.trailingComments, std::move(comments));
      }
    }

    // Track the same implicit-object/braced-object/array levels as the writer
    int CommentIndent(const ValuePath& path) const
    {
      const LosslessValue* current = &mRoot;
      std::size_t level = 0;
      bool arrayItem = false;
      std::size_t commentLevel = 0;
      for (const auto& segment : path)
      {
        if (const auto* key = std::get_if<std::string>(&segment))
        {
          if (!current->value.isObject())
          {
            return 0;
          }
          const auto index = MemberIndex(*current, *key);
          if (index == Append)
          {
            return 0;
          }
          level += arrayItem || !current->inlineComment.empty() || !current->closingComments.empty() ? 1 : 0;
          commentLevel = level;
          current = &current->objectItems[index].second;
          if (current->value.isObject() || current->value.isArray())
          {
            ++level;
          }
          arrayItem = false;
        }
        else
        {
          const auto index = std::get<std::size_t>(segment);
          if (!current->value.isArray() || index >= current->arrayItems.size())
          {
            return 0;
          }
          ++level;
          commentLevel = level;
          current = &current->arrayItems[index];
          arrayItem = true;
        }
      }
      const auto width = static_cast<std::size_t>(std::max(mOptions.indentWidth, 1));
      if (commentLevel > static_cast<std::size_t>(INT_MAX) / width)
      {
        return INT_MAX;
      }
      return static_cast<int>(commentLevel * width);
    }

    std::vector<LosslessComment> DetachedComments(const LosslessValue& node, int indent) const
    {
      auto comments = node.leadingComments;
      if (!node.blockComment.empty())
      {
        comments.push_back({indent, "#" + node.blockComment});
      }
      for (const auto& child : node.arrayItems)
      {
        AppendComments(comments, DetachedComments(child, indent));
      }
      for (const auto& member : node.objectItems)
      {
        AppendComments(comments, DetachedComments(member.second, indent));
      }
      AppendComments(comments, node.closingComments);
      if (!node.inlineComment.empty())
      {
        comments.push_back({indent, "#" + node.inlineComment});
      }
      AppendComments(comments, node.trailingComments);
      if (mOptions.commentIndentation == CommentIndentation::MatchDestination)
      {
        for (auto& comment : comments)
        {
          comment.indent = indent;
        }
      }
      return comments;
    }

    std::vector<LosslessComment> RemovedComments(const LosslessValue& node, int indent) const
    {
      return mOptions.removedComments == RemovedCommentPolicy::Preserve
        ? DetachedComments(node, indent) : std::vector<LosslessComment>{};
    }

    void PreserveReplacementComments(const LosslessValue& original, LosslessValue& replacement, int indent) const
    {
      const auto preserveLines = [](std::vector<LosslessComment>& target, const std::vector<LosslessComment>& source)
      {
        const bool same = target.size() == source.size() &&
          std::equal(target.begin(), target.end(), source.begin(),
            [](const auto& left, const auto& right) { return left.indent == right.indent && left.text == right.text; });
        if (!same)
        {
          Prepend(target, source);
        }
      };
      preserveLines(replacement.leadingComments, original.leadingComments);
      if (!original.blockComment.empty())
      {
        if (!replacement.blockComment.empty() && replacement.blockComment != original.blockComment)
        {
          replacement.leadingComments.push_back({indent, "#" + replacement.blockComment});
        }
        replacement.blockComment = original.blockComment;
      }
      if (!original.inlineComment.empty())
      {
        if (!replacement.inlineComment.empty() && replacement.inlineComment != original.inlineComment)
        {
          replacement.leadingComments.push_back({indent, "#" + replacement.inlineComment});
        }
        replacement.inlineComment = original.inlineComment;
      }
      std::vector<LosslessComment> nested;
      for (std::size_t index = 0; index < original.arrayItems.size(); ++index)
      {
        if (replacement.value.isArray() && index < replacement.arrayItems.size())
        {
          PreserveReplacementComments(original.arrayItems[index], replacement.arrayItems[index], indent);
        }
        else
        {
          AppendComments(nested, DetachedComments(original.arrayItems[index], indent));
        }
      }
      for (const auto& [key, child] : original.objectItems)
      {
        const auto index = replacement.value.isObject() ? MemberIndex(replacement, key) : Append;
        if (index != Append)
        {
          PreserveReplacementComments(child, replacement.objectItems[index].second, indent);
        }
        else
        {
          AppendComments(nested, DetachedComments(child, indent));
        }
      }
      preserveLines(replacement.closingComments, original.closingComments);
      Prepend(replacement.closingComments, std::move(nested));
      preserveLines(replacement.trailingComments, original.trailingComments);
      if (!replacement.value.isObject() && !replacement.value.isArray())
      {
        if (!replacement.blockComment.empty())
        {
          replacement.leadingComments.push_back({indent, "#" + replacement.blockComment});
          replacement.blockComment.clear();
        }
        Prepend(replacement.trailingComments, std::move(replacement.closingComments));
        replacement.closingComments.clear();
      }
    }
  };

  inline ErrorCode Parse(std::string_view src, Value& out, Error* error = nullptr, const ParseOptions& options = {})
  {
    Parser p(src, {}, options);
    return p.Parse(out, error);
  }

  namespace detail
  {
    // Lossless parser: preserves comment lines and ordering into LosslessValue
    class LosslessParser : private Parser
    {
    public:
      explicit LosslessParser(std::string_view src, std::string_view filename = {}, const ParseOptions& options = {})
        : Parser(src, filename, options)
      {}

      ErrorCode Parse(LosslessValue& out, Error* error)
      {
        LosslessValue parsed;
        const auto result = ParseDocument(parsed, error);
        if (result == ErrorCode::OK)
        {
          out = std::move(parsed);
        }
        return result;
      }

    private:
      ErrorCode ParseDocument(LosslessValue& out, Error* error)
      {
        ResetParseState();
        mPendingComments.clear();

        if (InputLimitExceeded())
        {
          return Finish(ErrorCode::ResourceLimit, error, "Maximum input byte count exceeded");
        }

        // Validate UTF-8 up front (strips leading BOM)
        std::size_t badIndex = 0;
        std::size_t badLine = 1;
        std::size_t badCol = 1;
        bool hasBOM = mSrc.size() >= 3 && static_cast<unsigned char>(mSrc[0]) == 0xEF && static_cast<unsigned char>(mSrc[1]) == 0xBB &&
          static_cast<unsigned char>(mSrc[2]) == 0xBF;
        if (!ValidateUTF8(mSrc, true, badIndex, badLine, badCol))
        {
          mPos = badIndex;
          mLine = badLine;
          mCol = badCol;
          return Finish(ErrorCode::InvalidUtf8, error, "Invalid UTF-8 encoding");
        }
        if (hasBOM)
        {
          mPos = 3;
          mCol = 1;
          mLine = 1;
        }

        bool hasLine = false;
        ErrorCode errorCode = NextContentLineLossless(hasLine, mPendingComments);
        if (errorCode != ErrorCode::OK)
        {
          return Finish(errorCode, error);
        }

        if (!hasLine)
        {
          return Finish(ErrorCode::UnexpectedEnd, error, "Empty document");
        }

        errorCode = ParseValueLossless(out, 0);
        if (errorCode != ErrorCode::OK)
        {
          return Finish(errorCode, error);
        }

        // Explicit top-level containers and scalars leave their final newline
        // to this caller. Preserve their inline and subsequent full-line trivia.
        if (!EndOfFile() && mCol != 1)
        {
          SkipInlineSpaces();
          if (Peek() == '#')
          {
            out.inlineComment = ReadInlineComment();
          }
          else if (Peek() == '\r' || Peek() == '\n')
          {
            const char newline = Get();
            if (newline == '\r' && Peek() == '\n')
            {
              Get();
            }
          }
          else if (!EndOfFile())
          {
            return Finish(ErrorCode::UnexpectedChar, error, "Trailing characters after top-level value");
          }
        }
        if (!EndOfFile())
        {
          hasLine = false;
          errorCode = NextContentLineLossless(hasLine, mPendingComments);
          if (errorCode != ErrorCode::OK)
          {
            return Finish(errorCode, error);
          }
          if (hasLine)
          {
            return Finish(ErrorCode::UnexpectedChar, error, "Trailing characters after top-level value");
          }
        }

        // Any remaining pending comments belong after the root value
        if (!mPendingComments.empty())
        {
          out.trailingComments.insert(out.trailingComments.end(), mPendingComments.begin(), mPendingComments.end());
          mPendingComments.clear();
        }

        // Reset indent stack to base level
        mIndentStack.clear();
        mIndentStack.push_back(0);

        if (!EndOfFile())
        {
          return Finish(ErrorCode::UnexpectedChar, error, "Trailing characters after top-level value");
        }

        if (error)
        {
          *error = {};
        }

        return ErrorCode::OK;
      }

    private:
      std::vector<LosslessComment> mPendingComments;

      std::string ReadInlineComment()
      {
        std::string comment;
        if (!Match('#'))
        {
          return comment;
        }
        while (!EndOfFile() && Peek() != '\r' && Peek() != '\n')
        {
          comment.push_back(Get());
        }
        SkipToEOL();
        return comment;
      }

      ErrorCode Finish(ErrorCode errorCode, Error* error, std::optional<std::string_view> message = std::nullopt)
      {
        // Prefer existing detailed error (e.g., from base Fail) unless a new message is supplied
        if (message.has_value())
        {
          mError.code = errorCode;
          mError.where = Location();
          mError.message.assign(message->data(), message->size());
          mError.filename.assign(mFilename.begin(), mFilename.end());
        }
        else if (mError.code == ErrorCode::OK)
        {
          mError.code = errorCode;
          mError.where = Location();
          mError.message.clear();
          mError.filename.assign(mFilename.begin(), mFilename.end());
        }

        if (error)
        {
          if (message.has_value())
          {
            *error = {};
            error->code = errorCode;
            error->where = Location();
            error->message.assign(message->data(), message->size());
            error->filename.assign(mFilename.begin(), mFilename.end());
          }
          else if (mError.code != ErrorCode::OK)
          {
            *error = mError;
          }
          else
          {
            *error = {};
            error->code = errorCode;
            error->where = Location();
            error->message.clear();
            error->filename.assign(mFilename.begin(), mFilename.end());
          }
        }
        return errorCode;
      }

      // Collect comment / blank lines into out while advancing to the next content line
      ErrorCode NextContentLineLossless(bool& hasLine, std::vector<LosslessComment>& comments)
      {
        hasLine = false;
        while (!EndOfFile())
        {
          if (mCol != 1)
          {
            SkipToEOL();
            if (EndOfFile())
            {
              return ErrorCode::OK;
            }
            continue;
          }

          int indentCols = 0;
          bool hasContent = false;
          std::size_t lineStart = mPos;
          ErrorCode errorCode = ReadLineIndent(indentCols, hasContent);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }

          if (!hasContent)
          {
            std::string line(mSrc.data() + lineStart, mPos - lineStart);
            while (!line.empty() && (line.back() == '\r' || line.back() == '\n'))
            {
              line.pop_back();
            }
            // Store indent and text after indent
            LosslessComment losslessComment;
            losslessComment.indent = indentCols;
            if (static_cast<std::size_t>(indentCols) < line.size())
            {
              losslessComment.text = line.substr(static_cast<std::size_t>(indentCols));
            }
            else
            {
              losslessComment.text.clear(); // Blank line
            }
            comments.push_back(std::move(losslessComment));
            continue;
          }

          errorCode = ApplyIndentLevel(indentCols);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }

          hasLine = true;
          return ErrorCode::OK;
        }
        return ErrorCode::OK; // EOF
      }

      ErrorCode ParseValueLossless(LosslessValue& out, int currentIndent, bool consumePending = true)
      {
        const auto begin = mPos;
        const auto result = ParseValueLosslessBody(out, currentIndent, consumePending);
        if (result == ErrorCode::OK)
        {
          RecordSource(out.value, begin);
        }
        return result;
      }

      ErrorCode ParseValueLosslessBody(LosslessValue& out, int currentIndent, bool consumePending)
      {
        if (consumePending && !mPendingComments.empty())
        {
          out.leadingComments.insert(out.leadingComments.end(), mPendingComments.begin(), mPendingComments.end());
          mPendingComments.clear();
        }

        char c = Peek();
        if (c == '{')
        {
          out.value = Object{};
          return ParseInlineObjectLossless(out, currentIndent);
        }
        if (c == '[')
        {
          out.value = Array{};
          return ParseArrayLossless(out, currentIndent);
        }
        if (c == '"' || c == '\'')
        {
          return ParseQuotedStringOrIndentedObjectLossless(out, currentIndent);
        }
        if (IsIdentifierStart(c))
        {
          return ParseIdentifierOrIndentedObjectLossless(out, currentIndent);
        }
        if (IsNumberStart(c))
        {
          Value tempValue;
          ErrorCode errorCode = ParseNumber(tempValue);
          if (errorCode == ErrorCode::OK)
          {
            out.value = std::move(tempValue);
          }
          return errorCode;
        }
        if (EndOfFile())
        {
          return Finish(ErrorCode::UnexpectedEnd, nullptr, "Unexpected end of input while parsing value");
        }
        return Finish(ErrorCode::UnexpectedChar, nullptr, "Unexpected character while parsing value");
      }

      ErrorCode ParseQuotedStringOrIndentedObjectLossless(LosslessValue& out, int currentIndent)
      {
        const std::size_t savedPos = mPos;
        const std::size_t savedLine = mLine;
        const std::size_t savedColumn = mCol;
        Value quotedValue;
        ErrorCode errorCode = Peek() == '"' ? ParseStringOrTriple(quotedValue) : ParseStringSingle(quotedValue);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        SkipInlineSpaces();
        if (Peek() != ':')
        {
          out.value = std::move(quotedValue);
          return ErrorCode::OK;
        }
        mPos = savedPos;
        mLine = savedLine;
        mCol = savedColumn;
        Object object;
        errorCode = ParseIndentedObjectBodyLossless(object, out, currentIndent);
        if (errorCode == ErrorCode::OK)
        {
          out.value = std::move(object);
        }
        return errorCode;
      }

      ErrorCode ParseInlineObjectLossless(LosslessValue& out, int currentIndent)
      {
        if (!CanEnterContainer())
        {
          return ErrorCode::ResourceLimit;
        }
        const ContainerDepthGuard depthGuard(mContainerDepth);
        if (!Match('{'))
        {
          return Finish(ErrorCode::InternalError, nullptr);
        }

        Object object;
        SkipInlineSpaces();

        const bool multiline = Peek() == '\r' || Peek() == '\n';
        if (multiline)
        {
          char newline = Get();
          if (newline == '\r' && Peek() == '\n')
          {
            Get();
          }
          bool hasLine = false;
          ErrorCode errorCode = NextContentLineLossless(hasLine, mPendingComments);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          if (!hasLine)
          {
            return Finish(ErrorCode::UnexpectedEnd, nullptr, "Unterminated inline object");
          }
          if (Peek() == '}')
          {
            out.closingComments.insert(out.closingComments.end(), mPendingComments.begin(), mPendingComments.end());
            mPendingComments.clear();
            Get();
            out.value = std::move(object);
            return ErrorCode::OK;
          }

          const int objectIndent = mIndentStack.back();
          if (objectIndent <= currentIndent)
          {
            return Finish(ErrorCode::InconsistentIndent, nullptr, "Expected indented members inside object braces");
          }

          while (true)
          {
            if (Peek() == '}')
            {
              out.closingComments.insert(out.closingComments.end(), mPendingComments.begin(), mPendingComments.end());
              mPendingComments.clear();
              Get();
              break;
            }

            std::vector<LosslessComment> preKeyComments;
            preKeyComments.swap(mPendingComments);
            const LocationEntry keyLocation = Location();
            std::string key;
            SourceSpan keySource;
            errorCode = ParseKey(key, &keySource);
            if (errorCode != ErrorCode::OK)
            {
              return errorCode;
            }
            errorCode = CheckDuplicateKey(object, key, keyLocation);
            if (errorCode != ErrorCode::OK)
            {
              return errorCode;
            }
            SkipInlineSpaces();
            if (!Match(':'))
            {
              return Finish(ErrorCode::UnexpectedChar, nullptr, "Expected ':' in inline object");
            }
            SkipInlineSpaces();

            LosslessValue child;
            child.leadingComments = std::move(preKeyComments);
            if (Peek() == '#')
            {
              child.blockComment = ReadInlineComment();
            }
            else if (Peek() == '\r' || Peek() == '\n')
            {
              newline = Get();
              if (newline == '\r' && Peek() == '\n')
              {
                Get();
              }
            }

            if (!child.blockComment.empty() || mCol == 1)
            {
              hasLine = false;
              errorCode = NextContentLineLossless(hasLine, mPendingComments);
              if (errorCode != ErrorCode::OK)
              {
                return errorCode;
              }
              if (!hasLine || mIndentStack.back() <= objectIndent)
              {
                return Finish(ErrorCode::InconsistentIndent, nullptr, "Expected an indented block value");
              }
              errorCode = ParseValueLossless(child, mIndentStack.back(), false);
            }
            else
            {
              errorCode = ParseValueLossless(child, objectIndent, false);
            }
            if (errorCode != ErrorCode::OK)
            {
              return errorCode;
            }

            RecordKeySource(child.value, keySource);
            out.objectItems.emplace_back(key, child);
            object.emplace(std::move(key), child.value);

            SkipInlineSpaces();
            bool commentConsumedNewline = false;
            if (Peek() == '#')
            {
              out.objectItems.back().second.inlineComment = ReadInlineComment();
              commentConsumedNewline = true;
            }
            if (!commentConsumedNewline && Match(','))
            {
              SkipInlineSpaces();
            }
            if (!commentConsumedNewline && (Peek() == '\r' || Peek() == '\n'))
            {
              newline = Get();
              if (newline == '\r' && Peek() == '\n')
              {
                Get();
              }
              commentConsumedNewline = true;
            }
            if (commentConsumedNewline)
            {
              hasLine = false;
              errorCode = NextContentLineLossless(hasLine, mPendingComments);
              if (errorCode != ErrorCode::OK)
              {
                return errorCode;
              }
              if (!hasLine)
              {
                return Finish(ErrorCode::UnexpectedEnd, nullptr, "Unterminated inline object");
              }
            }
            if (mIndentStack.back() < objectIndent && Peek() != '}')
            {
              return Finish(ErrorCode::InconsistentIndent, nullptr, "Expected closing '}' for inline object");
            }
          }

          out.value = std::move(object);
          return ErrorCode::OK;
        }

        SkipWhitespaceAndComments();

        if (Match('}'))
        {
          out.value = object;
          return ErrorCode::OK;
        }

        while (true)
        {
          // Comments collected before this key belong to this entry
          std::vector<LosslessComment> preKeyComments;
          if (!mPendingComments.empty())
          {
            preKeyComments.swap(mPendingComments);
          }

          const LocationEntry keyLocation = Location();
          std::string key;
          SourceSpan keySource;
          ErrorCode errorCode = ParseKey(key, &keySource);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          errorCode = CheckDuplicateKey(object, key, keyLocation);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          SkipWhitespaceAndComments();
          if (!Match(':'))
          {
            return Finish(ErrorCode::UnexpectedChar, nullptr, "Expected ':' in inline object");
          }
          SkipWhitespaceAndComments();
          LosslessValue child;
          errorCode = ParseValueLossless(child, currentIndent);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          RecordKeySource(child.value, keySource);
          out.objectItems.emplace_back(key, child);
          object.emplace(std::move(key), child.value);
          SkipWhitespaceAndComments();
          if (Match('}'))
          {
            break;
          }
          if (!Match(','))
          {
            return Finish(ErrorCode::UnexpectedChar, nullptr, "Expected ',' in inline object");
          }
          SkipWhitespaceAndComments();
        }
        out.value = std::move(object);
        return ErrorCode::OK;
      }

      // Match the inline array grammar's whitespace handling without discarding
      // trivia or changing the indentation stack. A comment after '[' does not
      // make an inline array newline-separated; commas are still required.
      void CollectArrayTriviaLossless(LosslessValue* previous = nullptr)
      {
        // Nested indentation parsing may already have advanced to a new line.
        // Only attach inline trivia if this physical line contains a token.
        bool hasContentOnLine = false;
        for (auto position = mPos; position > 0;)
        {
          const char character = mSrc[--position];
          if (character == '\r' || character == '\n')
          {
            break;
          }
          if (character != ' ' && character != '\t')
          {
            hasContentOnLine = true;
            break;
          }
        }
        bool sameLine = previous != nullptr && hasContentOnLine;
        bool blankLine = !hasContentOnLine;
        while (!EndOfFile())
        {
          SkipInlineSpaces();
          if (Peek() == '#')
          {
            const int indent = static_cast<int>(mCol) - 1;
            auto comment = ReadInlineComment();
            if (sameLine && previous && previous->inlineComment.empty() &&
                mPendingComments.empty() && !comment.empty())
            {
              previous->inlineComment = std::move(comment);
            }
            else
            {
              mPendingComments.push_back({indent, "#" + comment});
            }
            sameLine = false;
            blankLine = true;
          }
          else if (Peek() == '\r' || Peek() == '\n')
          {
            if (blankLine)
            {
              mPendingComments.push_back({0, {}});
            }
            SkipToEOL();
            sameLine = false;
            blankLine = true;
          }
          else
          {
            break;
          }
        }
      }

      ErrorCode ParseArrayLossless(LosslessValue& out, int parentIndent)
      {
        if (!CanEnterContainer())
        {
          return ErrorCode::ResourceLimit;
        }
        const ContainerDepthGuard depthGuard(mContainerDepth);
        if (!Match('['))
        {
          return Finish(ErrorCode::InternalError, nullptr);
        }

        Array array;
        std::optional<std::size_t> closingEnd;
        const auto closeArray = [&]()
        {
          if (!Match(']'))
          {
            return false;
          }
          out.closingComments.insert(out.closingComments.end(), mPendingComments.begin(), mPendingComments.end());
          mPendingComments.clear();
          closingEnd = mPos;
          return true;
        };

        SkipInlineSpaces();
        bool isMultiline = (Peek() == '\r' || Peek() == '\n');

        if (isMultiline)
        {
          char c = Get();
          if (c == '\r' && Peek() == '\n')
          {
            Get();
          }

          bool hasLine = false;
          ErrorCode errorCode = NextContentLineLossless(hasLine, mPendingComments);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          if (!hasLine || Peek() == ']')
          {
            if (Peek() == ']')
            {
              (void)closeArray();
            }
            out.value = std::move(array);
            mArrayClosingEnd = closingEnd;
            return ErrorCode::OK;
          }

          int arrayIndent = mIndentStack.back();
          while (true)
          {
            if (Peek() == ']')
            {
              (void)closeArray();
              break;
            }

            LosslessValue child;
            errorCode = ParseValueLossless(child, arrayIndent);
            if (errorCode != ErrorCode::OK)
            {
              return errorCode;
            }
            array.push_back(child.value);
            out.arrayItems.push_back(std::move(child));

            SkipInlineSpaces();
            if (Peek() == '#')
            {
              // If comment starts at current indent, treat as leading comment for next element
              if (mCol <= static_cast<std::size_t>(arrayIndent + 1))
              {
                std::string line;
                Get(); // consume '#'
                line.push_back('#');
                while (!EndOfFile())
                {
                  char ch = Peek();
                  if (ch == '\r' || ch == '\n')
                  {
                    break;
                  }
                  line.push_back(Get());
                }
                LosslessComment losslessComment;
                losslessComment.indent = static_cast<int>(mCol) - 1;
                if (losslessComment.indent < 0)
                {
                  losslessComment.indent = 0;
                }
                losslessComment.text = std::move(line);
                mPendingComments.push_back(std::move(losslessComment));
                SkipToEOL();
                bool hasLine2 = false;
                errorCode = NextContentLineLossless(hasLine2, mPendingComments);
                if (errorCode != ErrorCode::OK)
                {
                  return errorCode;
                }
                if (!hasLine2)
                {
                  break;
                }
                continue;
              }
              else
              {
                std::string comment;
                Get();
                while (!EndOfFile())
                {
                  char ch = Peek();
                  if (ch == '\r' || ch == '\n')
                  {
                    break;
                  }
                  comment.push_back(Get());
                }
                if (!out.arrayItems.empty())
                {
                  out.arrayItems.back().inlineComment = comment;
                }
                SkipToEOL();
                bool hasLine2 = false;
                errorCode = NextContentLineLossless(hasLine2, mPendingComments);
                if (errorCode != ErrorCode::OK)
                {
                  return errorCode;
                }
                if (!hasLine2)
                {
                  break;
                }
                continue;
              }
            }
            if (closeArray())
            {
              break;
            }
            if (Match(','))
            {
              CollectArrayTriviaLossless(&out.arrayItems.back());
              continue;
            }
            if (Peek() == '\r' || Peek() == '\n')
            {
              c = Get();
              if (c == '\r' && Peek() == '\n')
              {
                Get();
              }
              bool hasLine2 = false;
              errorCode = NextContentLineLossless(hasLine2, mPendingComments);
              if (errorCode != ErrorCode::OK)
              {
                return errorCode;
              }
              if (!hasLine2 || mIndentStack.back() < arrayIndent)
              {
                break;
              }
              continue;
            }
            return Finish(ErrorCode::UnexpectedChar, nullptr, "Expected ',' or ']' or newline in multiline array");
          }

          // After loop, consume closing ] if present
          if (Peek() == ']')
          {
            (void)closeArray();
          }
        }
        else
        {
          CollectArrayTriviaLossless();
          if (closeArray())
          {
            out.value = std::move(array);
            mArrayClosingEnd = closingEnd;
            return ErrorCode::OK;
          }

          while (true)
          {
            if (Peek() == ']')
            {
              (void)closeArray();
              break;
            }

            LosslessValue child;
            ErrorCode errorCode = ParseValueLossless(child, parentIndent);
            if (errorCode != ErrorCode::OK)
            {
              return errorCode;
            }
            array.push_back(child.value);
            out.arrayItems.push_back(std::move(child));

            CollectArrayTriviaLossless(&out.arrayItems.back());
            if (closeArray())
            {
              break;
            }
            if (!Match(','))
            {
              return Finish(ErrorCode::UnexpectedChar, nullptr, "Expected ',' or ']' in inline array");
            }
            CollectArrayTriviaLossless(&out.arrayItems.back());
          }
        }

        out.value = std::move(array);
        mArrayClosingEnd = closingEnd;
        return ErrorCode::OK;
      }

      ErrorCode ParseIdentifierOrIndentedObjectLossless(LosslessValue& out, int currentIndent)
      {
        const auto savedPos = mPos;
        const auto savedColumn = mCol;
        std::string identifier;
        while (IsIdentifierChar(Peek()))
        {
          identifier.push_back(Get());
        }
        SkipInlineSpaces();
        if (Peek() == ':')
        {
          mPos = savedPos;
          mCol = savedColumn;
          Object object;
          ErrorCode errorCode = ParseIndentedObjectBodyLossless(object, out, currentIndent);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          out.value = std::move(object);
          return ErrorCode::OK;
        }

        if (identifier == "true")
        {
          out.value = true;
          return ErrorCode::OK;
        }
        if (identifier == "false")
        {
          out.value = false;
          return ErrorCode::OK;
        }
        if (identifier == "null")
        {
          out.value = nullptr;
          return ErrorCode::OK;
        }
        out.value = identifier;
        return ErrorCode::OK;
      }

      ErrorCode ParseIndentedObjectBodyLossless(Object& obj, LosslessValue& outWrapper, int parentIndent)
      {
        if (!CanEnterContainer())
        {
          return ErrorCode::ResourceLimit;
        }
        const ContainerDepthGuard depthGuard(mContainerDepth);
        int bodyIndent = -1;

        while (true)
        {
          if (bodyIndent == -1)
          {
            bodyIndent = mIndentStack.back();
          }

          SkipInlineSpaces();
          if (EndOfFile())
          {
            break;
          }

          if (Peek() == '#' || Peek() == '\r' || Peek() == '\n')
          {
            bool hasLine = false;
            ErrorCode errorCode = NextContentLineLossless(hasLine, mPendingComments);
            if (errorCode != ErrorCode::OK)
            {
              return errorCode;
            }
            if (!hasLine)
            {
              break;
            }
            if (mIndentStack.back() < bodyIndent)
            {
              break;
            }
            continue;
          }

          char c = Peek();
          if (!IsIdentifierStart(c) && c != '"' && c != '\'')
          {
            break;
          }

          // Comments collected before this key belong to this entry
          std::vector<LosslessComment> preKeyComments;
          if (!mPendingComments.empty())
          {
            preKeyComments.swap(mPendingComments);
          }

          const LocationEntry keyLocation = Location();
          std::string key;
          SourceSpan keySource;
          ErrorCode errorCode = ParseKey(key, &keySource);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          errorCode = CheckDuplicateKey(obj, key, keyLocation);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }

          SkipInlineSpaces();
          if (!Match(':'))
          {
            return Finish(ErrorCode::UnexpectedChar, nullptr, "Expected ':' in object pair");
          }
          SkipInlineSpaces();

          bool blockValue = false;
          if (Peek() == '#')
          {
            blockValue = true;
          }
          else if (Peek() == '\r' || Peek() == '\n')
          {
            blockValue = true;
          }

          if (blockValue)
          {
            std::string blockComment;
            if (Peek() == '#')
            {
              blockComment = ReadInlineComment();
            }
            else
            {
              char c2 = Get();
              if (c2 == '\r' && Peek() == '\n')
              {
                Get();
              }
            }

            bool hasLine = false;
            ErrorCode errorCode2 = NextContentLineLossless(hasLine, mPendingComments);
            if (errorCode2 != ErrorCode::OK)
            {
              return errorCode2;
            }
            if (!hasLine)
            {
              return Finish(ErrorCode::InconsistentIndent, nullptr, "Expected indented block after ':'");
            }

            if (mIndentStack.back() <= bodyIndent)
            {
              return Finish(ErrorCode::InconsistentIndent, nullptr, "Expected deeper indentation for block value");
            }

            LosslessValue child;
            child.blockComment = std::move(blockComment);
            if (!preKeyComments.empty())
            {
              child.leadingComments.insert(child.leadingComments.end(), preKeyComments.begin(), preKeyComments.end());
              preKeyComments.clear();
            }
            errorCode2 = ParseValueLossless(child, mIndentStack.back(), false);
            if (errorCode2 != ErrorCode::OK)
            {
              return errorCode2;
            }

            RecordKeySource(child.value, keySource);
            outWrapper.objectItems.emplace_back(key, child);
            obj.emplace(std::move(key), child.value);

            // After parsing block value, check if we need to advance to next line
            SkipInlineSpaces();
            bool advanceToContent = false;
            if (Peek() == '#')
            {
              outWrapper.objectItems.back().second.inlineComment = ReadInlineComment();
              advanceToContent = true;
            }
            else if (Peek() == '\r' || Peek() == '\n')
            {
              char c2 = Get();
              if (c2 == '\r' && Peek() == '\n')
              {
                Get();
              }
              advanceToContent = true;
            }
            if (advanceToContent)
            {
              bool nextHasLine = false;
              ErrorCode ec3 = NextContentLineLossless(nextHasLine, mPendingComments);
              if (ec3 != ErrorCode::OK)
              {
                return ec3;
              }
              if (!nextHasLine)
              {
                break;
              }
            }

            if (mIndentStack.back() < bodyIndent)
            {
              break;
            }
            continue;
          }

          LosslessValue child;
          if (!preKeyComments.empty())
          {
            child.leadingComments.insert(child.leadingComments.end(), preKeyComments.begin(), preKeyComments.end());
            preKeyComments.clear();
          }
          errorCode = ParseValueLossless(child, parentIndent);
          if (errorCode != ErrorCode::OK)
          {
            return errorCode;
          }
          RecordKeySource(child.value, keySource);
          outWrapper.objectItems.emplace_back(key, child);
          obj.emplace(std::move(key), child.value);

          SkipInlineSpaces();
          if (Peek() == '#')
          {
            // If comment starts at current indent, treat as leading comment for next key
            if (mCol <= static_cast<std::size_t>(bodyIndent + 1))
            {
              std::string line;
              Get();
              line.push_back('#');
              while (!EndOfFile())
              {
                char ch = Peek();
                if (ch == '\r' || ch == '\n')
                {
                  break;
                }
                line.push_back(Get());
              }
              LosslessComment losslessComment;
              losslessComment.indent = static_cast<int>(mCol) - 1;
              if (losslessComment.indent < 0)
              {
                losslessComment.indent = 0;
              }
              losslessComment.text = std::move(line);
              mPendingComments.push_back(std::move(losslessComment));
              SkipToEOL();
              bool hasLine = false;
              ErrorCode errorCode2 = NextContentLineLossless(hasLine, mPendingComments);
              if (errorCode2 != ErrorCode::OK)
              {
                return errorCode2;
              }
              if (!hasLine)
              {
                break;
              }
              if (mIndentStack.back() < bodyIndent)
              {
                break;
              }
              continue;
            }
            else
            {
              std::string comment;
              Get();
              while (!EndOfFile())
              {
                char ch = Peek();
                if (ch == '\r' || ch == '\n')
                {
                  break;
                }
                comment.push_back(Get());
              }
              if (!outWrapper.objectItems.empty())
              {
                outWrapper.objectItems.back().second.inlineComment = comment;
              }
              SkipToEOL();
              bool hasLine = false;
              ErrorCode errorCode2 = NextContentLineLossless(hasLine, mPendingComments);
              if (errorCode2 != ErrorCode::OK)
              {
                return errorCode2;
              }
              if (!hasLine)
              {
                break;
              }
              if (mIndentStack.back() < bodyIndent)
              {
                break;
              }
              continue;
            }
          }
          if (Match(','))
          {
            SkipInlineSpaces();
            continue;
          }
          if (Peek() == '\r' || Peek() == '\n')
          {
            char c2 = Get();
            if (c2 == '\r' && Peek() == '\n')
            {
              Get();
            }

            bool hasLine = false;
            ErrorCode errorCode2 = NextContentLineLossless(hasLine, mPendingComments);
            if (errorCode2 != ErrorCode::OK)
            {
              return errorCode2;
            }
            if (!hasLine)
            {
              break;
            }
            if (mIndentStack.back() < bodyIndent)
            {
              break;
            }
            continue;
          }
          break;
        }

        return ErrorCode::OK;
      }
    };
  } // namespace detail

  inline ErrorCode ParseLossless(std::string_view src, LosslessValue& out, Error* error = nullptr, const ParseOptions& options = {})
  {
    detail::LosslessParser losslessParser(src, {}, options);
    return losslessParser.Parse(out, error);
  }

  namespace detail
  {
    inline bool SetFileError(Error* error, std::string_view message,
                             const std::string& path, std::string_view operation,
                             std::error_code cause)
    {
      if (error)
      {
        *error = {};
        error->code = ErrorCode::IoError;
        error->message = message;
        error->filename = path;
        error->operation = operation;
        error->systemError = cause;
      }
      return false;
    }

    inline std::error_code ErrnoCode(int value = errno)
    {
      return std::error_code(value, std::generic_category());
    }

    inline std::error_code NativeFileError()
    {
#ifdef _WIN32
      return std::error_code(static_cast<int>(GetLastError()), std::system_category());
#else
      return ErrnoCode();
#endif
    }

    inline ErrorCode ReadFileUTF8(const std::string& path, std::string& data, Error* error,
                                  std::size_t maxInputBytes = 0)
    {
      auto fileStream = OpenFileUTF8(path, "rb");
      if (!fileStream)
      {
        SetFileError(error, "Failed to open file", path, "open", ErrnoCode());
        return ErrorCode::IoError;
      }

      // Read incrementally instead of allocating a pre-statted size: the file
      // can grow between a size query and a read. One extra byte tests the cap.
      std::string result;
      std::array<char, 8192> buffer{};
      while (true)
      {
        std::size_t count = buffer.size();
        if (maxInputBytes != 0)
        {
          const auto remaining = maxInputBytes - result.size();
          if (remaining < count)
          {
            count = remaining + 1;
          }
        }
        errno = 0;
        const auto read = std::fread(buffer.data(), 1, count, fileStream.get());
        const auto readError = ErrnoCode();
        if (std::ferror(fileStream.get()))
        {
          SetFileError(error, "Failed to read file", path, "read", readError);
          return ErrorCode::IoError;
        }
        if (maxInputBytes != 0 && read > maxInputBytes - result.size())
        {
          if (error)
          {
            *error = {};
            error->code = ErrorCode::ResourceLimit;
            error->message = "Maximum input byte count exceeded";
            error->filename = path;
          }
          return ErrorCode::ResourceLimit;
        }
        result.append(buffer.data(), read);
        if (read < count)
        {
          break;
        }
      }
      if (std::fclose(fileStream.release()) != 0)
      {
        SetFileError(error, "Failed to close file", path, "close", ErrnoCode());
        return ErrorCode::IoError;
      }
      data = std::move(result);
      if (error)
      {
        *error = {};
      }
      return ErrorCode::OK;
    }
  } // namespace detail

  inline ErrorCode ParseFile(const std::string& path, Value& out, Error* error = nullptr, const ParseOptions& options = {})
  {
    std::string data;
    const ErrorCode readResult = detail::ReadFileUTF8(path, data, error, options.maxInputBytes);
    if (readResult != ErrorCode::OK)
    {
      return readResult;
    }
    Parser parser(std::string_view(data), path, options);
    return parser.Parse(out, error);
  }

  inline ErrorCode ParseFileLossless(const std::string& path, LosslessValue& out, Error* error = nullptr, const ParseOptions& options = {})
  {
    std::string data;
    const ErrorCode readResult = detail::ReadFileUTF8(path, data, error, options.maxInputBytes);
    if (readResult != ErrorCode::OK)
    {
      return readResult;
    }
    detail::LosslessParser losslessParser(std::string_view(data), path, options);
    return losslessParser.Parse(out, error);
  }

  // Exception type and throwing wrappers
  struct ParseException : std::exception
  {
    Error error;
    explicit ParseException(Error e) : error(std::move(e))
    {}
    const char* what() const noexcept override
    {
      return error.message.c_str();
    }
  };

  inline Value ParseOrThrow(std::string_view src, const ParseOptions& options = {})
  {
    Value value;
    Error error;
    ErrorCode errorCode = Parse(src, value, &error, options);
    if (errorCode != ErrorCode::OK)
    {
      throw ParseException(error);
    }
    return value;
  }

  inline Value ParseFileOrThrow(const std::string& path, const ParseOptions& options = {})
  {
    Value value;
    Error error;
    ErrorCode errorCode = ParseFile(path, value, &error, options);
    if (errorCode != ErrorCode::OK)
    {
      throw ParseException(error);
    }
    return value;
  }

  inline LosslessValue ParseLosslessOrThrow(std::string_view src, const ParseOptions& options = {})
  {
    LosslessValue value;
    Error error;
    const ErrorCode errorCode = ParseLossless(src, value, &error, options);
    if (errorCode != ErrorCode::OK)
    {
      throw ParseException(error);
    }
    return value;
  }

  struct WriteOptions
  {
    int indentWidth = 2; // Spaces per indent level
    bool sortObjectKeys = false;
  };

  namespace detail
  {
    inline void WriteIndent(std::string& out, int level, int width)
    {
      out.append(static_cast<std::size_t>(level * width), ' ');
    }

    enum class WriteContext : std::uint8_t
    {
      Root,
      InObject,
      InArray,
    };

    inline void WriteValue(const Value& value, std::string& out, int indentLevel, const WriteOptions& opt, WriteContext ctx);

    using ObjectItemView = std::pair<std::string_view, const Value*>;

    inline std::vector<ObjectItemView> OrderedObjectItems(const Object& object, bool sortKeys)
    {
      std::vector<ObjectItemView> items;
      items.reserve(object.size());
      for (const auto& keyValuePair : object)
      {
        items.emplace_back(keyValuePair.first, &keyValuePair.second);
      }
      if (sortKeys)
      {
        std::sort(items.begin(), items.end(), [](const auto& a, const auto& b) { return a.first < b.first; });
      }
      return items;
    }

    using LosslessItemView = std::pair<std::string_view, const LosslessValue*>;

    inline std::vector<LosslessItemView>
    OrderedLosslessItems(const std::vector<std::pair<std::string, LosslessValue>>& items, bool sortKeys)
    {
      std::vector<LosslessItemView> views;
      views.reserve(items.size());
      for (const auto& keyValuePair : items)
      {
        views.emplace_back(keyValuePair.first, &keyValuePair.second);
      }
      if (sortKeys)
      {
        std::sort(views.begin(), views.end(), [](const auto& a, const auto& b) { return a.first < b.first; });
      }
      return views;
    }

    inline void WriteStringQuoted(std::string_view value, std::string& out)
    {
      out.push_back('"');
      for (char c : value)
      {
        switch (c)
        {
          case '"': out += "\\\""; break;
          case '\\': out += "\\\\"; break;
          case '\n': out += "\\n"; break;
          case '\r': out += "\\r"; break;
          case '\t': out += "\\t"; break;
          default: out.push_back(c); break;
        }
      }
      out.push_back('"');
    }

    inline bool CanWriteKeyWithoutQuotes(std::string_view key)
    {
      const auto isStart = [](char c) {
        return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_';
      };
      const auto isContinue = [&](char c) {
        return isStart(c) || (c >= '0' && c <= '9') || c == '-';
      };
      if (key.empty() || !isStart(key.front()))
      {
        return false;
      }
      return std::all_of(key.begin(), key.end(), isContinue);
    }

    inline void WriteKey(std::string_view key, std::string& out)
    {
      if (CanWriteKeyWithoutQuotes(key))
      {
        out.append(key);
      }
      else
      {
        WriteStringQuoted(key, out);
      }
    }

    inline std::string FormatNumber(double value)
    {
      std::array<char, 128> buffer{};
      const auto result = std::to_chars(buffer.data(), buffer.data() + buffer.size(), value, std::chars_format::general);
      if (result.ec != std::errc{})
      {
        return {};
      }
      std::string text(buffer.data(), result.ptr);
      if (text.find_first_of(".eE") == std::string::npos)
      {
        text += ".0";
      }
      return text;
    }

    inline bool ValidateFiniteNumbers(const Value& value, Error* error)
    {
      if (std::holds_alternative<double>(value))
      {
        if (!std::isfinite(std::get<double>(value)))
        {
          if (error)
          {
            *error = {};
            error->code = ErrorCode::InvalidNumber;
            error->where = {};
            error->message = "Non-finite number value";
          }
          return false;
        }
        return true;
      }
      if (std::holds_alternative<Array>(value))
      {
        for (const auto& child : std::get<Array>(value))
        {
          if (!ValidateFiniteNumbers(child, error))
          {
            return false;
          }
        }
        return true;
      }
      if (std::holds_alternative<Object>(value))
      {
        for (const auto& entry : std::get<Object>(value))
        {
          if (!ValidateFiniteNumbers(entry.second, error))
          {
            return false;
          }
        }
        return true;
      }
      return true;
    }

    inline bool ValidateFiniteNumbers(const LosslessValue& value, Error* error)
    {
      if (std::holds_alternative<Array>(value.value) && !value.arrayItems.empty())
      {
        for (const auto& child : value.arrayItems)
        {
          if (!ValidateFiniteNumbers(child, error))
          {
            return false;
          }
        }
        return true;
      }
      if (std::holds_alternative<Object>(value.value) && !value.objectItems.empty())
      {
        for (const auto& entry : value.objectItems)
        {
          if (!ValidateFiniteNumbers(entry.second, error))
          {
            return false;
          }
        }
        return true;
      }
      return ValidateFiniteNumbers(value.value, error);
    }

    // "Simple" scalars for inline arrays
    inline bool IsSimpleScalar(const Value& value)
    {
      if (std::holds_alternative<std::nullptr_t>(value) || std::holds_alternative<bool>(value) || std::holds_alternative<double>(value))
      {
        return true;
      }
      if (std::holds_alternative<std::string>(value))
      {
        const auto& readonlyValue = std::get<std::string>(value);
        // Single-line, reasonably short
        return readonlyValue.find('\n') == std::string::npos && readonlyValue.size() <= 32;
      }
      return false; // Arrays / objects aren't "simple"
    }

    inline bool CanInlineArray(const Array& array)
    {
      if (array.size() > 3)
      {
        return false;
      }
      for (const auto& value : array)
      {
        if (!IsSimpleScalar(value))
        {
          return false;
        }
      }
      return true;
    }

    // Forward declaration so inline writer can recurse on objects
    inline void WriteObject(
      const Object& object,
      std::string& out,
      int indentLevel,
      const WriteOptions& options,
      WriteContext ctx,
      bool indentFirstLine);

    // Inline writer used for array-context objects to avoid newlines
    inline void WriteValueInline(const Value& value, std::string& out, const WriteOptions& options)
    {
      if (std::holds_alternative<std::nullptr_t>(value))
      {
        out += "null";
      }
      else if (std::holds_alternative<bool>(value))
      {
        out += (std::get<bool>(value) ? "true" : "false");
      }
      else if (std::holds_alternative<double>(value))
      {
        out += FormatNumber(std::get<double>(value));
      }
      else if (std::holds_alternative<std::string>(value))
      {
        WriteStringQuoted(std::get<std::string>(value), out);
      }
      else if (std::holds_alternative<Array>(value))
      {
        out.push_back('[');
        std::size_t index = 0;
        for (const auto& e : std::get<Array>(value))
        {
          if (index++ > 0)
          {
            out += ", ";
          }
          WriteValueInline(e, out, options);
        }
        out.push_back(']');
      }
      else
      {
        // Object
        WriteObject(std::get<Object>(value), out, 0, options, WriteContext::InArray, false);
      }
    }

    inline void WriteObject(
      const Object& object,
      std::string& out,
      int indentLevel,
      const WriteOptions& options,
      WriteContext ctx,
      bool indentFirstLine = true)
    {
      auto items = OrderedObjectItems(object, options.sortObjectKeys);

      // If writing an object as an array element, emit fully inline braces to avoid indent ambiguity
      if (ctx == WriteContext::InArray)
      {
        out.push_back('{');
        std::size_t index = 0;
        for (const auto& [key, valuePtr] : items)
        {
          const Value& value = *valuePtr;
          if (index++ > 0)
          {
            out += ", ";
          }
          WriteKey(key, out);
          out += ": ";
          WriteValueInline(value, out, options);
        }
        out.push_back('}');
        return;
      }

      // CoffeeScript-style brace-less objects elsewhere
      const int contentIndent = indentLevel;

      // Empty object
      if (object.empty())
      {
        if (indentFirstLine)
        {
          WriteIndent(out, indentLevel, options.indentWidth);
        }
        out += "{}";
        return;
      }

      bool first = true;
      for (const auto& [key, valuePtr] : items)
      {
        const Value& value = *valuePtr;
        if (!first)
        {
          out.push_back('\n');
        }

        // If caller already emitted indent for the first line, skip it once
        if (!(first && !indentFirstLine))
        {
          WriteIndent(out, contentIndent, options.indentWidth);
        }
        first = false;

        WriteKey(key, out);

        out += ": ";

        const bool isObject = std::holds_alternative<Object>(value);
        const bool isArray = std::holds_alternative<Array>(value);
        const bool inlineArray = isArray && CanInlineArray(std::get<Array>(value));

        if (isObject || (isArray && !inlineArray))
        {
          // Block value on next line with increased indent
          out.push_back('\n');
          WriteValue(value, out, contentIndent + 1, options, WriteContext::InObject);
        }
        else
        {
          // Scalars and small inline arrays stay on the same line
          WriteValue(value, out, contentIndent, options, WriteContext::InObject);
        }
      }
    }

    inline void WriteArray(const Array& array, std::string& out, int indentLevel, const WriteOptions& options, WriteContext ctx)
    {
      // LOGIC:
      //  1. If array is small and all "simple" scalars -> inline: [1, 2, 3]
      //  2. Otherwise -> multiline:
      //       [
      //         1
      //         2
      //         foo: 1
      //         bar: 2
      //       ]
      //     where objects become brace-less indent blocks

      if (array.empty())
      {
        out += "[]";
        return;
      }

      bool allSimple = true;
      for (const auto& value : array)
      {
        if (!IsSimpleScalar(value))
        {
          allSimple = false;
          break;
        }
      }

      // Inline small scalar arrays
      if (allSimple && array.size() <= 3)
      {
        out.push_back('[');
        bool first = true;
        for (const auto& value : array)
        {
          if (!first)
          {
            out += ", ";
          }
          first = false;
          WriteValue(value, out, indentLevel, options, WriteContext::InArray);
        }
        out.push_back(']');
        return;
      }

      // Multiline array, CoffeeScript-style
      const bool callerIndented = ctx == WriteContext::InArray;
      if (!callerIndented)
      {
        WriteIndent(out, indentLevel, options.indentWidth);
      }
      out.push_back('[');
      out.push_back('\n');

      bool first = true;
      for (const auto& value : array)
      {
        if (!first)
        {
          out.push_back('\n');
        }
        first = false;

        // Every element begins at the array's child indent
        WriteIndent(out, indentLevel + 1, options.indentWidth);

        if (std::holds_alternative<Object>(value))
        {
          const Object& object = std::get<Object>(value);
          // Use array-context writer (braced, comma-separated) and skip the first indent we already wrote
          WriteObject(object, out, indentLevel + 1, options, WriteContext::InArray, false);
        }
        else
        {
          // Scalars / nested arrays keep the usual "one per line" indent
          WriteValue(value, out, indentLevel + 1, options, WriteContext::InArray);
        }
      }

      out.push_back('\n');
      WriteIndent(out, indentLevel, options.indentWidth);
      out.push_back(']');
    }

    inline void WriteValue(const Value& value, std::string& out, int indentLevel, const WriteOptions& options, WriteContext ctx)
    {
      if (std::holds_alternative<std::nullptr_t>(value))
      {
        out += "null";
      }
      else if (std::holds_alternative<bool>(value))
      {
        out += (std::get<bool>(value) ? "true" : "false");
      }
      else if (std::holds_alternative<double>(value))
      {
        out += FormatNumber(std::get<double>(value));
      }
      else if (std::holds_alternative<std::string>(value))
      {
        const auto& stringValue = std::get<std::string>(value);
        if (stringValue.find('\n') != std::string::npos)
        {
          // Multiline -> triple-quoted
          out += "\"\"\"";
          out += stringValue;
          out += "\"\"\"";
        }
        else
        {
          WriteStringQuoted(stringValue, out);
        }
      }
      else if (std::holds_alternative<Array>(value))
      {
        // Array
        WriteArray(std::get<Array>(value), out, indentLevel, options, ctx);
      }
      else
      {
        // Object
        WriteObject(std::get<Object>(value), out, indentLevel, options, ctx);
      }
    }

  } // namespace detail

  inline bool ToString(const Value& value, std::string& out, const WriteOptions& options = {}, Error* error = nullptr)
  {
    if (!detail::ValidateFiniteNumbers(value, error))
    {
      return false;
    }
    out.clear();
    detail::WriteValue(value, out, 0, options, detail::WriteContext::Root);
    if (error)
    {
      *error = {};
    }
    return true;
  }

  inline std::string ToString(const Value& value, const WriteOptions& options = {})
  {
    std::string result;
    Error error;
    if (!ToString(value, result, options, &error))
    {
      return {};
    }
    return result;
  }

  namespace detail
  {
    inline void EnsureNewline(std::string& out)
    {
      if (!out.empty() && out.back() != '\n')
      {
        out.push_back('\n');
      }
    }

    inline void WriteCommentLines(const std::vector<LosslessComment>& lines, std::string& out)
    {
      for (const auto& line : lines)
      {
        if (!line.text.empty())
        {
          WriteIndent(out, std::max(line.indent, 0), 1);
          out.append(line.text);
        }
        out.push_back('\n');
      }
    }

    inline void WriteInlineComment(std::string_view comment, std::string& out)
    {
      while (!comment.empty() && (comment.back() == ' ' || comment.back() == '\t'))
      {
        comment.remove_suffix(1);
      }
      if (!comment.empty())
      {
        out.append(" #");
        out.append(comment);
      }
    }

    inline void WriteTrailingComments(const LosslessValue& value, std::string& out)
    {
      if (!value.trailingComments.empty())
      {
        EnsureNewline(out);
        WriteCommentLines(value.trailingComments, out);
      }
    }

    inline void WriteLosslessValue(
      const LosslessValue& value,
      std::string& out,
      int indentLevel,
      const WriteOptions& options,
      WriteContext ctx,
      bool emitLeadingComments = true);

    inline void WriteLosslessMembers(
      const std::vector<std::pair<std::string, LosslessValue>>& members,
      std::string& out,
      int indentLevel,
      const WriteOptions& options)
    {
      const auto items = OrderedLosslessItems(members, options.sortObjectKeys);
      bool first = true;
      for (const auto& [key, childPtr] : items)
      {
        const LosslessValue& child = *childPtr;
        if (!first)
        {
          EnsureNewline(out);
        }
        first = false;

        if (!child.leadingComments.empty())
        {
          WriteCommentLines(child.leadingComments, out);
        }
        WriteIndent(out, indentLevel, options.indentWidth);
        WriteKey(key, out);
        out.push_back(':');

        const bool isContainer = std::holds_alternative<Array>(child.value) || std::holds_alternative<Object>(child.value);
        if (isContainer)
        {
          WriteInlineComment(child.blockComment, out);
          out.push_back('\n');
          WriteLosslessValue(child, out, indentLevel + 1, options, WriteContext::InObject, false);
        }
        else
        {
          out.push_back(' ');
          WriteValue(child.value, out, indentLevel, options, WriteContext::InObject);
          WriteInlineComment(child.inlineComment, out);
          WriteTrailingComments(child, out);
        }
      }
    }

    inline void WriteLosslessValue(
      const LosslessValue& value,
      std::string& out,
      int indentLevel,
      const WriteOptions& options,
      WriteContext ctx,
      bool emitLeadingComments)
    {
      if (emitLeadingComments && !value.leadingComments.empty())
      {
        WriteCommentLines(value.leadingComments, out);
      }

      if (std::holds_alternative<Array>(value.value) &&
          (!value.arrayItems.empty() || (value.value.asArray().empty() && !value.closingComments.empty())))
      {
        WriteIndent(out, indentLevel, options.indentWidth);
        out.append("[\n");
        for (const auto& child : value.arrayItems)
        {
          WriteLosslessValue(child, out, indentLevel + 1, options, WriteContext::InArray);
          EnsureNewline(out);
        }
        WriteCommentLines(value.closingComments, out);
        WriteIndent(out, indentLevel, options.indentWidth);
        out.push_back(']');
        WriteInlineComment(value.inlineComment, out);
        WriteTrailingComments(value, out);
        return;
      }

      if (std::holds_alternative<Object>(value.value) &&
          (!value.objectItems.empty() || (value.value.asObject().empty() && !value.closingComments.empty())))
      {
        const bool useBraces =
          ctx == WriteContext::InArray || !value.inlineComment.empty() || !value.closingComments.empty();
        if (useBraces)
        {
          WriteIndent(out, indentLevel, options.indentWidth);
          out.append("{\n");
          WriteLosslessMembers(value.objectItems, out, indentLevel + 1, options);
          EnsureNewline(out);
          WriteCommentLines(value.closingComments, out);
          WriteIndent(out, indentLevel, options.indentWidth);
          out.push_back('}');
          WriteInlineComment(value.inlineComment, out);
        }
        else
        {
          WriteLosslessMembers(value.objectItems, out, indentLevel, options);
        }
        WriteTrailingComments(value, out);
        return;
      }

      const bool isObjectInArray = ctx == WriteContext::InArray && std::holds_alternative<Object>(value.value);
      if (isObjectInArray)
      {
        WriteIndent(out, indentLevel, options.indentWidth);
        WriteValue(value.value, out, indentLevel, options, ctx);
      }
      else if (std::holds_alternative<Object>(value.value))
      {
        WriteValue(value.value, out, indentLevel, options, ctx);
      }
      else if (std::holds_alternative<Array>(value.value))
      {
        WriteIndent(out, indentLevel, options.indentWidth);
        WriteValue(value.value, out, indentLevel, options, WriteContext::InArray);
      }
      else
      {
        WriteIndent(out, indentLevel, options.indentWidth);
        WriteValue(value.value, out, indentLevel, options, ctx);
      }
      WriteInlineComment(value.inlineComment, out);
      WriteTrailingComments(value, out);
    }
  } // namespace detail

  inline bool ToStringLossless(const LosslessValue& value, std::string& out, const WriteOptions& options = {}, Error* error = nullptr)
  {
    if (!detail::ValidateFiniteNumbers(value, error))
    {
      return false;
    }
    out.clear();
    detail::WriteLosslessValue(value, out, 0, options, detail::WriteContext::Root);
    if (error)
    {
      *error = {};
    }
    return true;
  }

  inline std::string ToStringLossless(const LosslessValue& value, const WriteOptions& options = {})
  {
    std::string result;
    Error error;
    if (!ToStringLossless(value, result, options, &error))
    {
      return {};
    }
    return result;
  }

  namespace detail
  {
    inline bool WriteFileContents(const std::string& path, std::string_view contents, Error* error)
    {
      auto fileStream = OpenFileUTF8(path, "wb");
      if (!fileStream)
      {
        return SetFileError(error, "Failed to open file for writing", path, "open", ErrnoCode());
      }

      std::size_t offset = 0;
      while (offset < contents.size())
      {
        errno = 0;
        const std::size_t written = std::fwrite(contents.data() + offset, 1, contents.size() - offset, fileStream.get());
        if (written == 0)
        {
          return SetFileError(error, "Failed to write file", path, "write", ErrnoCode());
        }
        offset += written;
      }
      if (std::fflush(fileStream.get()) != 0)
      {
        return SetFileError(error, "Failed to flush file", path, "flush", ErrnoCode());
      }

      std::FILE* rawFile = fileStream.release();
      if (std::fclose(rawFile) != 0)
      {
        return SetFileError(error, "Failed to close file", path, "close", ErrnoCode());
      }
      if (error)
      {
        *error = {};
      }
      return true;
    }
  } // namespace detail

  inline bool WriteFile(const std::string& path, const Value& value, const WriteOptions& options = {}, Error* error = nullptr)
  {
    std::string stringValue;
    if (!ToString(value, stringValue, options, error))
    {
      return false;
    }
    return detail::WriteFileContents(path, stringValue, error);
  }

  inline bool WriteFileLossless(const std::string& path, const LosslessValue& value, const WriteOptions& options = {}, Error* error = nullptr)
  {
    std::string stringValue;
    if (!ToStringLossless(value, stringValue, options, error))
    {
      return false;
    }
    return detail::WriteFileContents(path, stringValue, error);
  }

#ifdef HAVCSON_ENABLE_TEST_HOOKS
  // Test-only atomic-write failure injection
  namespace testing
  {
    enum class AtomicWriteStage : std::uint8_t
    {
      None,
      TemporaryOpen,
      Write,
      Flush,
      Sync,
      Close,
      Replace,
    };

    inline thread_local AtomicWriteStage AtomicWriteFailureStage = AtomicWriteStage::None;

    inline void FailAtomicWriteAt(AtomicWriteStage stage)
    {
      AtomicWriteFailureStage = stage;
    }
  } // namespace testing
#endif

  namespace detail
  {
#ifdef HAVCSON_ENABLE_TEST_HOOKS
    inline bool ShouldFailAtomicWrite(testing::AtomicWriteStage stage)
    {
      return testing::AtomicWriteFailureStage == stage;
    }
#endif

    inline std::string AtomicTemporaryPath(
      const std::string& destination,
      std::uint64_t uniqueSequence,
      unsigned int attempt)
    {
#ifdef _WIN32
      const auto processId = static_cast<unsigned long long>(GetCurrentProcessId());
#else
      const auto processId = static_cast<unsigned long long>(::getpid());
#endif
      return destination + ".tmp." + std::to_string(processId) + "." +
             std::to_string(uniqueSequence) + "." + std::to_string(attempt);
    }

#ifdef _WIN32
    inline bool WriteTextFileAtomicImpl(const std::string& path, std::string_view contents, Error* error)
    {
      static volatile LONG nextUniqueSequence = 0;
      const std::uint64_t uniqueSequence = static_cast<std::uint32_t>(InterlockedIncrement(&nextUniqueSequence));
      std::wstring temporaryPath;
      HANDLE file = INVALID_HANDLE_VALUE;

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::TemporaryOpen))
      {
        return SetFileError(error, "Simulated atomic temporary-file open failure", path, "open", std::make_error_code(std::errc::io_error));
      }
#endif

      for (unsigned int attempt = 0; attempt < 128; ++attempt)
      {
        temporaryPath = ConvertStringToWString(AtomicTemporaryPath(path, uniqueSequence, attempt), true);
        if (temporaryPath.empty())
        {
          return SetFileError(error, "Invalid UTF-8 atomic temporary-file path", path, "open", std::make_error_code(std::errc::illegal_byte_sequence));
        }
        file = CreateFileW(temporaryPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file != INVALID_HANDLE_VALUE)
        {
          break;
        }
        const DWORD openError = GetLastError();
        if (openError != ERROR_FILE_EXISTS && openError != ERROR_ALREADY_EXISTS)
        {
          return SetFileError(error, "Failed to create atomic temporary file", path, "open", std::error_code(static_cast<int>(openError), std::system_category()));
        }
      }
      if (file == INVALID_HANDLE_VALUE)
      {
        return SetFileError(error, "Failed to allocate a unique atomic temporary file", path, "open", std::make_error_code(std::errc::file_exists));
      }

      auto fail = [&](std::string_view message, std::string_view operation, std::error_code cause) {
        if (file != INVALID_HANDLE_VALUE)
        {
          CloseHandle(file);
          file = INVALID_HANDLE_VALUE;
        }
        DeleteFileW(temporaryPath.c_str());
        return SetFileError(error, message, path, operation, cause);
      };

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Write))
      {
        return fail("Simulated atomic write failure", "write", std::make_error_code(std::errc::io_error));
      }
#endif

      std::size_t offset = 0;

      while (offset < contents.size())
      {
        const std::size_t remaining = contents.size() - offset;
        const DWORD chunk = static_cast<DWORD>(std::min(remaining, static_cast<std::size_t>(std::numeric_limits<DWORD>::max())));
        DWORD written = 0;
        if (!::WriteFile(file, contents.data() + offset, chunk, &written, nullptr))
        {
          return fail("Failed to write atomic temporary file", "write", NativeFileError());
        }
        if (written == 0)
        {
          return fail("Atomic write made no progress", "write", {});
        }
        offset += static_cast<std::size_t>(written);
      }

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Flush))
      {
        return fail("Simulated atomic flush failure", "flush", std::make_error_code(std::errc::io_error));
      }
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Sync))
      {
        return fail("Simulated atomic sync failure", "sync", std::make_error_code(std::errc::io_error));
      }
#endif
      if (!FlushFileBuffers(file))
      {
        return fail("Failed to flush atomic temporary file", "flush", NativeFileError());
      }

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Close))
      {
        return fail("Simulated atomic close failure", "close", std::make_error_code(std::errc::io_error));
      }
#endif
      if (!CloseHandle(file))
      {
        const auto cause = NativeFileError();
        file = INVALID_HANDLE_VALUE;
        return fail("Failed to close atomic temporary file", "close", cause);
      }
      file = INVALID_HANDLE_VALUE;

      const std::wstring destinationPath = ConvertStringToWString(path, true);
      if (destinationPath.empty())
      {
        return fail("Invalid UTF-8 destination path", "replace", std::make_error_code(std::errc::illegal_byte_sequence));
      }
#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Replace))
      {
        return fail("Simulated atomic replacement failure", "replace", std::make_error_code(std::errc::io_error));
      }
#endif
      if (!MoveFileExW(
            temporaryPath.c_str(), destinationPath.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
      {
        return fail("Failed to atomically replace destination file", "replace", NativeFileError());
      }
      if (error)
      {
        *error = {};
      }
      return true;
    }
#else
    inline bool SyncContainingDirectory(const std::string& path, Error* error)
    {
      const std::size_t slash = path.find_last_of('/');
      const std::string directory = slash == std::string::npos ? "." : (slash == 0 ? "/" : path.substr(0, slash));
      int flags = O_RDONLY;
#ifdef O_DIRECTORY
      flags |= O_DIRECTORY;
#endif
#ifdef O_CLOEXEC
      flags |= O_CLOEXEC;
#endif
      const int directoryFile = ::open(directory.c_str(), flags);
      if (directoryFile < 0)
      {
        return SetFileError(error, "Atomic replacement succeeded, but its directory could not be opened for sync", path, "openDirectory", ErrnoCode());
      }
      const bool synced = ::fsync(directoryFile) == 0;
      const auto syncError = ErrnoCode();
      const bool closed = ::close(directoryFile) == 0;
      if (!synced || !closed)
      {
        return SetFileError(error, "Atomic replacement succeeded, but its directory could not be synced", path,
                            synced ? "closeDirectory" : "syncDirectory", synced ? ErrnoCode() : syncError);
      }
      return true;
    }

    inline bool WriteTextFileAtomicImpl(const std::string& path, std::string_view contents, Error* error)
    {
      static std::atomic<std::uint64_t> nextUniqueSequence{0};
      const std::uint64_t uniqueSequence = nextUniqueSequence.fetch_add(1, std::memory_order_relaxed);
      std::string temporaryPath;
      int file = -1;

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::TemporaryOpen))
      {
        return SetFileError(error, "Simulated atomic temporary-file open failure", path, "open", std::make_error_code(std::errc::io_error));
      }
#endif

      for (unsigned int attempt = 0; attempt < 128; ++attempt)
      {
        temporaryPath = AtomicTemporaryPath(path, uniqueSequence, attempt);
        int flags = O_WRONLY | O_CREAT | O_EXCL;
#ifdef O_CLOEXEC
        flags |= O_CLOEXEC;
#endif
        file = ::open(temporaryPath.c_str(), flags, static_cast<mode_t>(0666));
        if (file >= 0)
        {
          break;
        }
        if (errno != EEXIST)
        {
          return SetFileError(error, "Failed to create atomic temporary file", path, "open", ErrnoCode());
        }
      }
      if (file < 0)
      {
        return SetFileError(error, "Failed to allocate a unique atomic temporary file", path, "open", std::make_error_code(std::errc::file_exists));
      }

      auto fail = [&](std::string_view message, std::string_view operation, std::error_code cause) {
        if (file >= 0)
        {
          ::close(file);
          file = -1;
        }
        ::unlink(temporaryPath.c_str());
        return SetFileError(error, message, path, operation, cause);
      };

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Write))
      {
        return fail("Simulated atomic write failure", "write", std::make_error_code(std::errc::io_error));
      }
#endif

      std::size_t offset = 0;

      while (offset < contents.size())
      {
        const std::size_t chunk = std::min(contents.size() - offset, static_cast<std::size_t>(std::numeric_limits<ssize_t>::max()));
        const ssize_t written = ::write(file, contents.data() + offset, chunk);
        if (written < 0 && errno == EINTR)
        {
          continue;
        }
        if (written < 0)
        {
          return fail("Failed to write atomic temporary file", "write", ErrnoCode());
        }
        if (written == 0)
        {
          return fail("Atomic write made no progress", "write", {});
        }
        offset += static_cast<std::size_t>(written);
      }

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Flush))
      {
        return fail("Simulated atomic flush failure", "flush", std::make_error_code(std::errc::io_error));
      }
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Sync))
      {
        return fail("Simulated atomic sync failure", "sync", std::make_error_code(std::errc::io_error));
      }
#endif
      if (::fsync(file) != 0)
      {
        return fail("Failed to sync atomic temporary file", "sync", ErrnoCode());
      }

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Close))
      {
        return fail("Simulated atomic close failure", "close", std::make_error_code(std::errc::io_error));
      }
#endif
      if (::close(file) != 0)
      {
        const auto cause = ErrnoCode();
        file = -1;
        return fail("Failed to close atomic temporary file", "close", cause);
      }

      file = -1;

#ifdef HAVCSON_ENABLE_TEST_HOOKS
      if (ShouldFailAtomicWrite(testing::AtomicWriteStage::Replace))
      {
        return fail("Simulated atomic replacement failure", "replace", std::make_error_code(std::errc::io_error));
      }
#endif
      if (::rename(temporaryPath.c_str(), path.c_str()) != 0)
      {
        return fail("Failed to atomically replace destination file", "replace", ErrnoCode());
      }
      if (!SyncContainingDirectory(path, error))
      {
        return false;
      }
      if (error)
      {
        *error = {};
      }
      return true;
    }
#endif
  } // namespace detail

  // Writes the serialized text to a temporary file in the destination directory,
  // flushes it to disk, and then replaces the destination file in one operation.
  inline bool WriteTextFileAtomic(const std::string& path, std::string_view contents, Error* error = nullptr)
  {
    if (path.empty() || path.find('\0') != std::string::npos)
    {
      return detail::SetFileError(error, "Destination path is empty or contains a null byte", path, "open",
                                  std::make_error_code(std::errc::invalid_argument));
    }
    return detail::WriteTextFileAtomicImpl(path, contents, error);
  }

  inline bool WriteFileAtomic(
    const std::string& path,
    const Value& value,
    const WriteOptions& options = {},
    Error* error = nullptr)
  {
    std::string serialized;
    if (!ToString(value, serialized, options, error))
    {
      return false;
    }
    return WriteTextFileAtomic(path, serialized, error);
  }

  inline bool WriteFileLosslessAtomic(
    const std::string& path,
    const LosslessValue& value,
    const WriteOptions& options = {},
    Error* error = nullptr)
  {
    std::string serialized;
    if (!ToStringLossless(value, serialized, options, error))
    {
      return false;
    }
    return WriteTextFileAtomic(path, serialized, error);
  }

  // Simple JSON writer without pretty-printing
  inline bool ToJsonString(const Value& value, std::string& out, Error* error = nullptr)
  {
    if (!detail::ValidateFiniteNumbers(value, error))
    {
      return false;
    }
    out.clear();

    // JSON writer
    struct Writer
    {
      static void JsonString(const std::string& value, std::string& out)
      {
        out.push_back('"');
        for (char c : value)
        {
          switch (c)
          {
            case '"': out += "\\\""; break;
            case '\\': out += "\\\\"; break;
            case '\n': out += "\\n"; break;
            case '\r': out += "\\r"; break;
            case '\t': out += "\\t"; break;
            default: out.push_back(c); break;
          }
        }
        out.push_back('"');
      }

      static void Write(const Value& value, std::string& out)
      {
        if (std::holds_alternative<std::nullptr_t>(value))
        {
          out += "null";
        }
        else if (std::holds_alternative<bool>(value))
        {
          out += (std::get<bool>(value) ? "true" : "false");
        }
        else if (std::holds_alternative<double>(value))
        {
          out += detail::FormatNumber(std::get<double>(value));
        }
        else if (std::holds_alternative<std::string>(value))
        {
          JsonString(std::get<std::string>(value), out);
        }
        else if (std::holds_alternative<Array>(value))
        {
          out.push_back('[');
          bool first = true;
          for (auto& e : std::get<Array>(value))
          {
            if (!first)
            {
              out.push_back(',');
            }
            first = false;
            Write(e, out);
          }
          out.push_back(']');
        }
        else
        {
          const Object& object = std::get<Object>(value);
          out.push_back('{');
          bool first = true;
          for (const auto& [key, childValue] : object)
          {
            if (!first)
            {
              out.push_back(',');
            }
            first = false;
            JsonString(key, out);
            out.push_back(':');
            Write(childValue, out);
          }
          out.push_back('}');
        }
      }
    };

    Writer::Write(value, out);
    if (error)
    {
      *error = {};
    }
    return true;
  }

  inline std::string ToJsonString(const Value& value)
  {
    std::string result;
    Error error;
    if (!ToJsonString(value, result, &error))
    {
      return {};
    }
    return result;
  }
}

#endif
