---
name: paginated-report
description: Edit Power BI paginated report files (.rdl) safely - add or format a table, header, footer, parameters or page settings - and check the XML against the RDL rules before handing it back. Use when asked to change, format or fix a paginated report or an .rdl file, or when Report Builder shows a "Deserialization failed" error.
---

# Paginated report (.rdl) editor

An .rdl file is XML in the RDL 2016 schema (`http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition`). Report Builder refuses the whole file if one element is wrong, so follow these rules exactly.

## Rules

- **Never touch the data source or dataset connection.** Create those in Report Builder (Power BI semantic model connection). Only change layout, formatting, parameters and expressions.
- **Back up first:** copy `<name>.rdl` to `<name>.backup.rdl` before the first edit.
- **Use only elements the schema allows.** Borders are written with a `<Border>` element inside `<Style>`, never `<BorderStyle>`:

  ```xml
  <Style>
    <Border>
      <Color>#cfd4da</Color>
      <Style>Solid</Style>
      <Width>1pt</Width>
    </Border>
  </Style>
  ```

  Allowed children of `<Style>` include Border, TopBorder, BottomBorder, LeftBorder, RightBorder, BackgroundColor, FontStyle, FontFamily, FontSize, FontWeight, Format, TextDecoration, TextAlign, VerticalAlign, Color, PaddingLeft, PaddingRight, PaddingTop, PaddingBottom (from Report Builder's own error message).
- Element names are case-sensitive. Keep every element in the order Report Builder wrote it. Don't invent elements.
- Field references use the dataset's field names exactly, e.g. `=Fields!Supplier.Value`.

## Steps

1. Read the .rdl and list what's in it: data sources, datasets and their fields, parameters, the body items.
2. Make a backup copy.
3. Make the requested change, following the rules.
4. Check: the file is well-formed XML, every `<Style>` child is in the allowed list, and every `Fields!X` exists in the dataset.
5. Tell the user what changed, and to reopen the file in Report Builder. If Report Builder shows an error, ask them to paste it: the message names the bad element, the expected elements, and the line and position.

## Never

- Publish, render or export the report. That happens in Report Builder or the Power BI service.
- Put credentials or connection strings in the file.
