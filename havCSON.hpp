/*
havCSON.hpp

ABOUT

Havoc's single-file CSON (CoffeeScript Object Notation) library for C++.

REVISION HISTORY

v0.1 (2025-12-15) - First release.

LICENSE

MIT License

Copyright (c) 2025 René Nicolaus

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
  #define _UNICODE
  #undef UNICODE
  #define UNICODE

  #undef NOMINMAX
  #define NOMINMAX

  #undef STRICT
  #define STRICT

  #ifndef _WIN32_WINNT
    #define _WIN32_WINNT _WIN32_WINNT_WINXP
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
#include <algorithm>
#include <exception>
#include <fstream>
#include <memory>
#include <optional>
#include <sstream>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <variant>
#include <vector>

static_assert(sizeof(signed char) == 1, "expected char to be 1 byte");
static_assert(sizeof(unsigned char) == 1, "expected unsigned char to be 1 byte");
static_assert(sizeof(signed char) == 1, "expected int8 to be 1 byte");
static_assert(sizeof(unsigned char) == 1, "expected uint8 to be 1 byte");
static_assert(sizeof(signed short int) == 2, "expected int16 to be 2 bytes");
static_assert(sizeof(unsigned short int) == 2, "expected uint16 to be 2 bytes");
static_assert(sizeof(signed int) == 4, "expected int32 to be 4 bytes");
static_assert(sizeof(unsigned int) == 4, "expected uint32 to be 4 bytes");
static_assert(sizeof(signed long long) == 8, "expected int64 to be 8 bytes");
static_assert(sizeof(unsigned long long) == 8, "expected uint64 to be 8 bytes");

namespace havCSON
{
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
  inline std::unique_ptr<std::FILE, decltype(&std::fclose)>
  OpenFileUTF8(const std::string& path, const std::string& mode)
  {
    std::unique_ptr<std::FILE, decltype(&std::fclose)> fileStream(nullptr, &std::fclose);
    std::wstring modeW = ConvertStringToWString(mode, true);
    std::wstring pathW = ConvertStringToWString(path, true);
    std::FILE* file = nullptr;
    if (_wfopen_s(&file, pathW.c_str(), modeW.c_str()) == 0)
    {
      fileStream.reset(file);
    }
    return fileStream;
  }
#else
  inline std::unique_ptr<std::FILE, decltype(&std::fclose)>
  OpenFileUTF8(const std::string& path, const std::string& mode)
  {
    std::unique_ptr<std::FILE, decltype(&std::fclose)> fileStream(nullptr, &std::fclose);
    fileStream.reset(std::fopen(path.c_str(), mode.c_str()));
    return fileStream;
  }
#endif

  struct LocationEntry
  {
    std::size_t line = 1;
    std::size_t column = 1;
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
  };

  struct Error
  {
    ErrorCode code = ErrorCode::OK;
    LocationEntry where{};
    std::string message;

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
  };

  class Parser
  {
  public:
    Parser(std::string_view src, std::string_view filename = {}) : mSrc(src), mFilename(filename)
    {}

    ErrorCode Parse(Value& out, Error* error = nullptr)
    {
      // Reset previous error state for a fresh parse
      mError = {};

      // Validate UTF-8 up front (strips leading BOM)
      std::size_t badIndex = 0;
      std::size_t badLine = 1;
      std::size_t badCol = 1;
      bool hasBOM = mSrc.size() >= 3 && static_cast<unsigned char>(mSrc[0]) == 0xEF &&
        static_cast<unsigned char>(mSrc[1]) == 0xBB && static_cast<unsigned char>(mSrc[2]) == 0xBF;
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
              error->code = errorCode;
              error->where = Location();
              error->message.clear();
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
    std::string_view mFilename;
    std::size_t mPos = 0;
    std::size_t mLine = 1;
    std::size_t mCol = 1;

    // CoffeeScript-like indent model
    int mIndentUnit = 0; // Discovered on first non-zero indent
    std::vector<int> mIndentStack{0}; // Known indent levels (columns)

    Error mError;

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
      if (c == '\n')
      {
        ++mLine;
        mCol = 1;
      }
      else
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
      if (out)
      {
        *out = mError;
      }
      return code;
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
        if (c == '\n')
        {
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
          c = Peek();
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

    bool ValidateUTF8(
      std::string_view stringView,
      bool allowLeadingBOM,
      std::size_t& badIndex,
      std::size_t& badLine,
      std::size_t& badCol)
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
          index == 0 && allowLeadingBOM && stringView.size() >= 3 && c == 0xEF &&
          static_cast<unsigned char>(stringView[1]) == 0xBB && static_cast<unsigned char>(stringView[2]) == 0xBF)
        {
          index += 3;
          continue;
        }
        else if (
          index > 0 && c == 0xEF && index + 2 < stringView.size() &&
          static_cast<unsigned char>(stringView[index + 1]) == 0xBB &&
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
        int length = 0;
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
          (length == 4 && (codePoint < 0x10000 || codePoint > 0x10FFFF)) ||
          (codePoint >= 0xD800 && codePoint <= 0xDFFF))
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
      char c = Peek();
      if (c == '{')
      {
        return ParseInlineObject(out, currentIndent);
      }
      if (c == '[')
      {
        return ParseArray(out, currentIndent);
      }
      if (c == '"')
      {
        // Could be normal or triple string
        return ParseStringOrTriple(out);
      }
      if (c == '\'')
      {
        // Single-quoted string (no multiline)
        return ParseStringSingle(out);
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
      std::string ident;
      LocationEntry startLoc = Location();
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
        mPos -= ident.size();
        mCol -= ident.size();
        Object object;
        ErrorCode errorCode = ParseIndentedObjectBody(object, currentIndent);
        if (errorCode != ErrorCode::OK)
        {
          return errorCode;
        }
        out = std::move(object);
        return ErrorCode::OK;
      }

      // Otherwise interpret as bare identifier value
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
      out = ident; // Bare string
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
      char* endPtr = nullptr;
      std::string tempValue(stringView.begin(), stringView.end());
      double value = std::strtod(tempValue.c_str(), &endPtr);
      if (!endPtr || *endPtr != '\0')
      {
        return Fail(ErrorCode::InvalidNumber, nullptr, "Invalid number literal");
      }
      out = value;
      return ErrorCode::OK;
    }

    ErrorCode ParseInlineObject(Value& out, int currentIndent)
    {
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
        std::string key;
        ErrorCode errorCode = ParseKey(key);
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

        std::string key;
        ErrorCode errorCode = ParseKey(key);
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
      if (!Match('['))
      {
        return Fail(ErrorCode::InternalError, nullptr);
      }

      Array array;

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
            Get();
          }
          out = std::move(array);
          return ErrorCode::OK;
        }

        // Parse multiline array elements
        int arrayIndent = mIndentStack.back();
        while (true)
        {
          if (Peek() == ']')
          {
            Get();
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
          if (Match(']'))
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
          Get();
        }
      }
      else
      {
        // Inline array
        SkipWhitespaceAndComments();
        if (Match(']'))
        {
          out = std::move(array);
          return ErrorCode::OK;
        }

        while (true)
        {
          if (Peek() == ']')
          {
            Get();
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
          if (Match(']'))
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
      return ErrorCode::OK;
    }

    ErrorCode ParseKey(std::string& outKey)
    {
      SkipInlineSpaces();
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

  inline ErrorCode Parse(std::string_view src, Value& out, Error* error = nullptr)
  {
    Parser p(src);
    return p.Parse(out, error);
  }

  namespace detail
  {
    // Lossless parser: preserves comment lines and ordering into LosslessValue
    class LosslessParser : private Parser
    {
    public:
      explicit LosslessParser(std::string_view src) : Parser(src)
      {}

      ErrorCode Parse(LosslessValue& out, Error* error)
      {
        // Reset previous error state for a fresh parse
        mError = {};

        // Validate UTF-8 up front (strips leading BOM)
        std::size_t badIndex = 0;
        std::size_t badLine = 1;
        std::size_t badCol = 1;
        bool hasBOM = mSrc.size() >= 3 && static_cast<unsigned char>(mSrc[0]) == 0xEF &&
          static_cast<unsigned char>(mSrc[1]) == 0xBB && static_cast<unsigned char>(mSrc[2]) == 0xBF;
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

      ErrorCode Finish(ErrorCode errorCode, Error* error, std::optional<std::string_view> message = std::nullopt)
      {
        // Prefer existing detailed error (e.g., from base Fail) unless a new message is supplied
        if (message.has_value())
        {
          mError.code = errorCode;
          mError.where = Location();
          mError.message.assign(message->data(), message->size());
        }
        else if (mError.code == ErrorCode::OK)
        {
          mError.code = errorCode;
          mError.where = Location();
          mError.message.clear();
        }

        if (error)
        {
          if (message.has_value())
          {
            error->code = errorCode;
            error->where = Location();
            error->message.assign(message->data(), message->size());
          }
          else if (mError.code != ErrorCode::OK)
          {
            *error = mError;
          }
          else
          {
            error->code = errorCode;
            error->where = Location();
            error->message.clear();
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
        if (c == '"')
        {
          Value tempValue;
          ErrorCode errorCode = ParseStringOrTriple(tempValue);
          if (errorCode == ErrorCode::OK)
          {
            out.value = std::move(tempValue);
          }
          return errorCode;
        }
        if (c == '\'')
        {
          Value tempValue;
          ErrorCode errorCode = ParseStringSingle(tempValue);
          if (errorCode == ErrorCode::OK)
          {
            out.value = std::move(tempValue);
          }
          return errorCode;
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

      ErrorCode ParseInlineObjectLossless(LosslessValue& out, int currentIndent)
      {
        if (!Match('{'))
        {
          return Finish(ErrorCode::InternalError, nullptr);
        }
        Object object;
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

          std::string key;
          ErrorCode errorCode = ParseKey(key);
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

      ErrorCode ParseArrayLossless(LosslessValue& out, int parentIndent)
      {
        if (!Match('['))
        {
          return Finish(ErrorCode::InternalError, nullptr);
        }

        Array array;

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
              Get();
            }
            out.value = std::move(array);
            return ErrorCode::OK;
          }

          int arrayIndent = mIndentStack.back();
          while (true)
          {
            if (Peek() == ']')
            {
              Get();
              break;
            }

            LosslessValue child;
            errorCode = ParseValueLossless(child, arrayIndent, false);
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
              if (mCol <= arrayIndent + 1)
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
            if (Match(']'))
            {
              break;
            }
            if (Match(','))
            {
              SkipWhitespaceAndComments();
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
            Get();
          }
        }
        else
        {
          SkipWhitespaceAndComments();
          if (Match(']'))
          {
            out.value = std::move(array);
            return ErrorCode::OK;
          }

          while (true)
          {
            if (Peek() == ']')
            {
              Get();
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

            SkipWhitespaceAndComments();
            if (Match(']'))
            {
              break;
            }
            if (!Match(','))
            {
              return Finish(ErrorCode::UnexpectedChar, nullptr, "Expected ',' or ']' in inline array");
            }
            SkipWhitespaceAndComments();
          }
        }

        out.value = std::move(array);
        return ErrorCode::OK;
      }

      ErrorCode ParseIdentifierOrIndentedObjectLossless(LosslessValue& out, int currentIndent)
      {
        std::string identifier;
        while (IsIdentifierChar(Peek()))
        {
          identifier.push_back(Get());
        }
        SkipInlineSpaces();
        if (Peek() == ':')
        {
          mPos -= identifier.size();
          mCol -= identifier.size();
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

          std::string key;
          ErrorCode errorCode = ParseKey(key);
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
            if (Peek() == '#')
            {
              SkipToEOL();
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

            outWrapper.objectItems.emplace_back(key, child);
            obj.emplace(std::move(key), child.value);

            // After parsing block value, check if we need to advance to next line
            SkipInlineSpaces();
            if (Peek() == '#')
            {
              SkipToEOL();
            }
            if (Peek() == '\r' || Peek() == '\n')
            {
              char c2 = Get();
              if (c2 == '\r' && Peek() == '\n')
              {
                Get();
              }

              bool hasLine = false;
              ErrorCode ec3 = NextContentLineLossless(hasLine, mPendingComments);
              if (ec3 != ErrorCode::OK)
              {
                return ec3;
              }
              if (!hasLine)
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
          outWrapper.objectItems.emplace_back(key, child);
          obj.emplace(std::move(key), child.value);

          SkipInlineSpaces();
          if (Peek() == '#')
          {
            // If comment starts at current indent, treat as leading comment for next key
            if (mCol <= bodyIndent + 1)
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

  inline ErrorCode ParseLossless(std::string_view src, LosslessValue& out, Error* error = nullptr)
  {
    detail::LosslessParser losslessParser(src);
    return losslessParser.Parse(out, error);
  }

  inline ErrorCode ParseFile(const std::string& path, Value& out, Error* error = nullptr)
  {
    auto fileStream = OpenFileUTF8(path, "rb");
    if (!fileStream)
    {
      if (error)
      {
        error->code = ErrorCode::InternalError;
        error->where = {};
        error->message = "Failed to open file";
      }
      return ErrorCode::InternalError;
    }
    if (std::fseek(fileStream.get(), 0, SEEK_END) != 0)
    {
      if (error)
      {
        error->code = ErrorCode::InternalError;
        error->where = {};
        error->message = "Failed to read file";
      }
      return ErrorCode::InternalError;
    }
    long size = std::ftell(fileStream.get());
    if (size < 0 || std::fseek(fileStream.get(), 0, SEEK_SET) != 0)
    {
      if (error)
      {
        error->code = ErrorCode::InternalError;
        error->where = {};
        error->message = "Failed to read file";
      }
      return ErrorCode::InternalError;
    }
    std::string data(static_cast<std::size_t>(size), '\0');
    if (!data.empty())
    {
      if (std::fread(&data[0], 1, static_cast<std::size_t>(size), fileStream.get()) != static_cast<std::size_t>(size))
      {
        if (error)
        {
          error->code = ErrorCode::InternalError;
          error->where = {};
          error->message = "Failed to read file";
        }
        return ErrorCode::InternalError;
      }
    }
    return Parse(std::string_view(data), out, error);
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

  inline Value ParseOrThrow(std::string_view src)
  {
    Value value;
    Error error;
    ErrorCode errorCode = Parse(src, value, &error);
    if (errorCode != ErrorCode::OK)
    {
      throw ParseException(error);
    }
    return value;
  }

  inline Value ParseFileOrThrow(const std::string& path)
  {
    Value value;
    Error error;
    ErrorCode errorCode = ParseFile(path, value, &error);
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

    inline void
    WriteValue(const Value& value, std::string& out, int indentLevel, const WriteOptions& opt, WriteContext ctx);

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

    // "Simple" scalars for inline arrays
    inline bool IsSimpleScalar(const Value& value)
    {
      if (
        std::holds_alternative<std::nullptr_t>(value) || std::holds_alternative<bool>(value) ||
        std::holds_alternative<double>(value))
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
        out += std::to_string(std::get<double>(value));
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
          bool bareOK = true;
          if (key.empty())
          {
            bareOK = false;
          }
          else
          {
            char c0 = key[0];
            if (!((c0 >= 'A' && c0 <= 'Z') || (c0 >= 'a' && c0 <= 'z') || c0 == '_'))
            {
              bareOK = false;
            }
            for (char c : key)
            {
              if (!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '-'))
              {
                bareOK = false;
                break;
              }
            }
          }
          if (bareOK)
          {
            out += key;
          }
          else
          {
            WriteStringQuoted(key, out);
          }
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

        // Key
        bool bareOK = true;
        if (key.empty())
        {
          bareOK = false;
        }
        else
        {
          char c0 = key[0];
          if (!((c0 >= 'A' && c0 <= 'Z') || (c0 >= 'a' && c0 <= 'z') || c0 == '_'))
          {
            bareOK = false;
          }
          for (char c : key)
          {
            if (!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '-'))
            {
              bareOK = false;
              break;
            }
          }
        }

        if (bareOK)
        {
          out += key;
        }
        else
        {
          WriteStringQuoted(key, out);
        }

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

    inline void
    WriteArray(const Array& array, std::string& out, int indentLevel, const WriteOptions& options, WriteContext ctx)
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

    inline void
    WriteValue(const Value& value, std::string& out, int indentLevel, const WriteOptions& options, WriteContext ctx)
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
        out += std::to_string(std::get<double>(value));
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

  inline std::string ToString(const Value& value, const WriteOptions& options = {})
  {
    std::string out;
    detail::WriteValue(value, out, 0, options, detail::WriteContext::Root);
    return out;
  }

  namespace detail
  {
    inline void
    WriteCommentLines(const std::vector<LosslessComment>& lines, std::string& out, const WriteOptions& options)
    {
      bool prevNonEmpty = false;
      for (const auto& line : lines)
      {
        if (prevNonEmpty && !line.text.empty())
        {
          out.push_back('\n');
        }
        if (line.text.empty())
        {
          out.push_back('\n');
          prevNonEmpty = false;
        }
        else
        {
          WriteIndent(out, line.indent, 1); // Indent is absolute columns
          out.append(line.text);
          out.push_back('\n');
          prevNonEmpty = true;
        }
      }
    }

    inline void WriteLosslessValue(
      const LosslessValue& value,
      std::string& out,
      int indentLevel,
      const WriteOptions& options,
      WriteContext ctx)
    {
      WriteCommentLines(value.leadingComments, out, options);

      auto trimTrailingSpaces = [](std::string& value) {
        while (!value.empty() && (value.back() == ' ' || value.back() == '\t'))
        {
          value.pop_back();
        }
      };

      auto writeInlineComment = [&](const std::string& value) {
        if (!value.empty())
        {
          out.append(" #");
          out.append(value);
        }
      };

      if (std::holds_alternative<Array>(value.value) && !value.arrayItems.empty())
      {
        WriteIndent(out, indentLevel, options.indentWidth);
        out.append("[\n");
        for (std::size_t index = 0; index < value.arrayItems.size(); ++index)
        {
          const LosslessValue& child = value.arrayItems[index];
          WriteLosslessValue(child, out, indentLevel + 1, options, WriteContext::InArray);
          if (index + 1 < value.arrayItems.size())
          {
            out.push_back('\n');
          }
        }
        out.push_back('\n');
        WriteIndent(out, indentLevel, options.indentWidth);
        out.push_back(']');
        return;
      }

      if (std::holds_alternative<Object>(value.value) && !value.objectItems.empty())
      {
        auto items = OrderedLosslessItems(value.objectItems, options.sortObjectKeys);
        // Use braces only in array context, otherwise use indent-style
        if (ctx == WriteContext::InArray)
        {
          // Inline braced format for objects in arrays
          out.append("{\n");
          for (std::size_t index = 0; index < items.size(); ++index)
          {
            const auto& [key, childPtr] = items[index];
            const LosslessValue& child = *childPtr;
            WriteCommentLines(child.leadingComments, out, options);
            WriteIndent(out, indentLevel + 1, options.indentWidth);
            out.append(key);
            out.append(": ");
            WriteValue(child.value, out, indentLevel + 1, options, WriteContext::InObject);
            writeInlineComment(child.inlineComment);
            if (index + 1 < items.size())
            {
              out.push_back('\n');
            }
          }
          out.push_back('\n');
          WriteIndent(out, indentLevel, options.indentWidth);
          out.push_back('}');
        }
        else
        {
          // Indent-style format for regular objects
          for (std::size_t index = 0; index < items.size(); ++index)
          {
            const auto& [key, childPtr] = items[index];
            const LosslessValue& child = *childPtr;

            // Write leading comments for this key (includes blanks)
            WriteCommentLines(child.leadingComments, out, options);

            WriteIndent(out, indentLevel, options.indentWidth);
            out.append(key);
            out.append(":");

            // Determine if value should be on next line
            bool isObject = std::holds_alternative<Object>(child.value);
            bool isArray = std::holds_alternative<Array>(child.value);

            if (isObject || isArray)
            {
              out.push_back('\n');
              // For nested objects / arrays, don't write their leading comments again (they were already written above
              // as comments for this key)
              std::vector<LosslessComment> savedComments = child.leadingComments;
              const_cast<LosslessValue&>(child).leadingComments.clear();
              WriteLosslessValue(child, out, indentLevel + 1, options, WriteContext::InObject);
              const_cast<LosslessValue&>(child).leadingComments = std::move(savedComments);
            }
            else
            {
              out.push_back(' ');
              WriteValue(child.value, out, indentLevel, options, WriteContext::InObject);
            }

            // Trim trailing spaces from inline comments
            std::string comment = child.inlineComment;
            trimTrailingSpaces(comment);
            writeInlineComment(comment);

            if (index + 1 < items.size())
            {
              out.push_back('\n');
            }
          }
        }
        return;
      }

      WriteIndent(out, indentLevel, options.indentWidth);
      WriteValue(value.value, out, indentLevel, options, ctx);
      writeInlineComment(value.inlineComment);
    }
  } // namespace detail

  inline std::string ToStringLossless(const LosslessValue& value, const WriteOptions& options = {})
  {
    std::string out;
    detail::WriteLosslessValue(value, out, 0, options, detail::WriteContext::Root);
    return out;
  }

  inline bool
  WriteFile(const std::string& path, const Value& value, const WriteOptions& options = {}, Error* error = nullptr)
  {
    auto fileStream = OpenFileUTF8(path, "wb");
    if (!fileStream)
    {
      if (error)
      {
        error->code = ErrorCode::InternalError;
        error->where = {};
        error->message = "Failed to open file for writing";
      }
      return false;
    }
    std::string stringValue = ToString(value, options);
    if (!stringValue.empty())
    {
      if (std::fwrite(stringValue.data(), 1, stringValue.size(), fileStream.get()) != stringValue.size())
      {
        if (error)
        {
          error->code = ErrorCode::InternalError;
          error->where = {};
          error->message = "Failed to write file";
        }
        return false;
      }
    }
    if (error)
    {
      *error = {};
    }
    return true;
  }

  inline bool WriteFileLossless(
    const std::string& path,
    const LosslessValue& value,
    const WriteOptions& options = {},
    Error* error = nullptr)
  {
    auto fileStream = OpenFileUTF8(path, "wb");
    if (!fileStream)
    {
      if (error)
      {
        error->code = ErrorCode::InternalError;
        error->where = {};
        error->message = "Failed to open file for writing";
      }
      return false;
    }
    std::string stringValue = ToStringLossless(value, options);
    if (!stringValue.empty())
    {
      if (std::fwrite(stringValue.data(), 1, stringValue.size(), fileStream.get()) != stringValue.size())
      {
        if (error)
        {
          error->code = ErrorCode::InternalError;
          error->where = {};
          error->message = "Failed to write file";
        }
        return false;
      }
    }
    if (error)
    {
      *error = {};
    }
    return true;
  }

  // Simple JSON writer without pretty-printing
  inline std::string ToJsonString(const Value& value)
  {
    std::string out;

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
          out += std::to_string(std::get<double>(value));
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
          for (auto& [key, value] : object)
          {
            if (!first)
            {
              out.push_back(',');
            }
            first = false;
            JsonString(key, out);
            out.push_back(':');
            Write(value, out);
          }
          out.push_back('}');
        }
      }
    };

    Writer::Write(value, out);
    return out;
  }
}

#endif
