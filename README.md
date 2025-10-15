# TypeID for Excel - Complete Implementation

TypeID implementation for Microsoft Excel with VBA, supporting UUIDv7 generation and base32 (Crockford) encoding.



## Quick Start

1. Download `TypeID.bas`
2. Open Excel and press `Alt+F11`
3. Go to `File` > `Import File` and select the `.bas` file
4. Use in your worksheet: `=GenerateTypeID("user")`

## Key Features

✅ **UUIDv7 generation** - Time-sortable unique identifiers
✅ **Base32 Crockford encoding** - Human-readable, URL-safe
✅ **130-bit encoding** - Fully compatible with TypeID specification
✅ **Anti-regeneration tools** - Batch generation, value freezing macros
✅ **Prefix validation** - Ensures TypeID compliance

## Main Functions

| Function | Description |
|----------|-------------|
| `GenerateTypeID(prefix)` | Generate a single TypeID (recalculates) |
| `GenerateTypeIDsBatch()` | Batch generation with UI |
| `InsertTypeID(prefix)` | Insert TypeID as value (no recalculation) |
| `ConvertFormulasToValues()` | Freeze existing formulas |
| `TypeIDIf(condition, prefix)` | Conditional generation |

## Example Usage

```vba
' Simple generation
=GenerateTypeID("user")
' Result: user_01h2xcejqtf2nbrexx3vqjhp41

' Conditional generation
=TypeIDIf(B2, "order")
' Result: Generates only if B2 is TRUE

' Batch generation (select cells, run macro)
GenerateTypeIDsBatch
```

## TypeID Format

```
prefix_26characteridentifier

Example: user_01h455vb4pex5vsknk084sn02q
         └──┘ └──────────────────────────┘
       prefix    26-char base32 encoded
                 UUIDv7 (130 bits)
```

## Technical Details

- **UUID Version**: v7 (time-sortable)
- **Encoding**: Base32 Crockford
- **Alphabet**: `0123456789abcdefghjkmnpqrstvwxyz`
- **Suffix Length**: 26 characters
- **Bits Encoded**: 130 (2 zero-padding + 128 UUID bits)
- **Prefix Rules**: lowercase a-z and underscore, max 63 chars

## Preventing Regeneration

The `GenerateTypeID()` function recalculates on sheet changes. To prevent this:

**Option 1: Batch Generation (Recommended)**
```
1. Select range
2. Run GenerateTypeIDsBatch macro
3. IDs are written as values (won't change)
```

**Option 2: Manual Copy-Paste**
```
1. Generate with formula
2. Copy cells (Ctrl+C)
3. Paste Special > Values (Ctrl+Alt+V, V, Enter)
```

**Option 3: Use Conditional Function**
```
=TypeIDIf(TRUE, "user")
' Will generate once and stay stable
```

## Validation

Validate generated TypeIDs with:
- Official TypeID library: https://github.com/jetify-com/typeid
- NPM package: https://www.npmjs.com/package/typeid-js
- Online decoder tools

Example validation in Node.js:
```javascript
const { TypeID } = require('typeid-js');
const tid = TypeID.fromString("user_01h2xcejqtf2nbrexx3vqjhp41");
console.log(tid.getType());      // "user"
console.log(tid.toUUID());       // UUID format
console.log(tid.getTimestamp()); // Timestamp in ms
```

## Compatibility

✅ Excel 2010+
✅ Excel for Windows
✅ Excel for Mac (with VBA support)

## Common Issues

**"Stringified UUID is invalid"**
- Solution: Use the latest version (supports 130-bit encoding)

**IDs regenerate on recalculation**
- Solution: Use batch generation or convert formulas to values

**"Invalid prefix" error**
- Solution: Use only lowercase a-z and underscore, max 63 chars

## License

This implementation follows the TypeID specification:
https://github.com/jetify-com/typeid/tree/main/spec


## Changelog

**v1.0** - Initial release with 130-bit encoding support
- UUIDv7 generation
- Base32 Crockford encoding
- Anti-regeneration tools
- Batch generation macros
- Full TypeID spec compliance
