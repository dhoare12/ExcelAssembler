# 📊 ExcelAssembler

A lightweight templating engine for Excel, using familiar `<Content Select="..."/>` and `<Repeat Select="..."/>` tags inspired by OpenXmlPowerTools' `DocumentAssembler`.

## ✨ Project Objectives

- Allow non-technical users to build Excel templates using simple XPath-driven tags
- Replace template fields with XML data at runtime
- Support repeated data (rows or sections) using `<Repeat>` blocks
- Preserve Excel-based formatting using a convention-driven approach
- Integrate seamlessly with existing workflows that already use DocumentAssembler

---

## ✅ Features Implemented

### 1. **Value Injection**
- Replace cells containing `<Content Select="..."/>` with values from an XML file
- Supports relative XPath (`./Policy/TotalPremium` etc.)
- Detects numeric formatting and parses accordingly (e.g. handles `£123.45` cleanly)

### 2. **Preserved Formatting**
- Template tags are placed in a row
- A second row beneath provides Excel formatting (e.g. currency, dates)
- Values are inserted into the format row and the tag row is removed during processing

### 3. **Repeat Blocks**
- Supports row-based repeat sections with:
  ```xml
  <Repeat Select="./Proposal/Products/Product" />
  <Content Select="./Name" />
  <Content Select="./Premium" />
  <EndRepeat />
- Repeats the template row for each matching XML element
- Values inserted into pre-formatted rows
- Entire repeat block is removed after population

## How to Use

1. Create a .xlsx template with tags:

    `<Content Select="..." />` for fields

   `<Repeat Select="..." />` and `<EndRepeat />` to define loops

2. Add formatting rows beneath <Content> rows
3. Save as Template.xlsx
4. Provide an XML data file (e.g. data.xml)
5. Run `new ExcelAssembler().ProcessTemplate("Template.xlsx", "data.xml", "Output.xlsx");`

## Roadmap / Next ideas
- Support nested <Repeat> blocks
- Add <If Select="..."> / <EndIf /> for conditional sections
- Support column-based or transposed repeats
- Add support for inline format hints (Format="currency")
- Build an Excel plugin UI (insert tags, preview XML structure, validate template)
- Logging/reporting for unresolved XPaths and validation

## Credit
Inspired by the OpenXmlPowerTools DocumentAssembler syntax — bringing the same approach to Excel.
