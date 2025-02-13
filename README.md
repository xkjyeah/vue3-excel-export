# vue3-excel-export

## Quick usage instructions

```html
<template>
  <SheetJsOutput>
    <!-- Define how we want the spreadsheet presented -->
    <template v-slot:sheet>
      <row>
        <text>Initial</text>
        <text>Name</text>
        <text>Birthday</text>
        <text>Age this year</text>
        <text>Age this year (formula)</text>
      </row>

      <row v-for="row in dataset" :widthSetting="true">
        <text width="2">{{ row['name'][0] }}</text>
        <text>{{ row['name'] }}</text>
        <date z="MMM DD" width="7">{{ row['birthday'] }}</date>
        <number>{{ today.getUTCFullYear() - new Date(row['birthday']).getUTCFullYear() }}</number>
        <formula>DATEDIF(<rc c="-2" />, NOW(), "Y")</formula>
      </row>
    </template>

    <template v-slot:default="sheetObject">
      <button @click="renderExcelAndDownload(sheetObject)">Download Excel</button>
    </template>
  </SheetJsOutput>
</template>

<script setup>
import XLSX from 'xlsx'
import SheetJsOutput from 'vue3-excel-export/SheetJsOutput'

const today = new Date

const dataset = [
  { name: 'Alan', birthday: '1999-01-02' },
  { name: 'Bob', birthday: '2000-03-04' },
]

function renderExcelAndDownload(worksheet) {
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, "demo.xlsx");
}
</script>
```

Result:

![image](https://github.com/user-attachments/assets/6a948a97-2e5d-4dba-aaf5-40c73ba542cb)



## Quick API documentation

You will notice that these are **not custom elements** (custom elements have a hyphen in the name, or use PascalCase naming),
This is illegal in HTML and non-idiomatic in Vue. However, this is intentional, since these elements
are actually being rendered to a spreadsheet and never to the DOM. In fact, I used a custom renderer
to implement SheetJsOutput.

### `<row [:widthSetting=true|false]>`
Inserts a new Excel row. If a row is width-setting, we will set the column widths in the sheet based on the
widths of the cells in this row.

### `<text [z=formatString] [width=n]>`, `<number [z=formatString] [width=n]>`, `<boolean [z=formatString] [width=n]>`
Inserts a new cell of string, number or boolean type respectively. The *entire text content* will be interpreted as a string, number or boolean.
Booleans are `true` only if the string evaluates to "true".

`width=n` is honoured if the cell is specified inside a row that's `widthSetting=true`. Width is specified in [MDW units](https://docs.sheetjs.com/docs/csf/features/colprops#column-widths).

### `<date [z=formatString] [width=n]>`
Inserts a new cell of number type. The text contents of the cell are interpreted as a Javascript date (i.e. `new Date(textContents)`),
and then converted to [Excel epoch](https://docs.sheetjs.com/docs/csf/features/dates), representing the number of days since 1899-12-30.

### `<formula [z=formatString] [t=typeCode]>` `<rc [r=n] [c=n] />`
Inserts a new cell with a formula, e.g. `<formula>A1 + B1</formula>`. Formula is specified in [A1 format](https://docs.sheetjs.com/docs/csf/general#a1-style).

It would obviously be immensely helpful if we could use R1C1 format, but SheetJS doesn't support that :(

Instead, you can use the `<rc>` tag to help:

```html
<row>
  <number>Qty</number>
  <number>Unit price</number>
  <number>Price</number>
</row>
<row>
  <number>{{ item.qty }}</number>
  <number>{{ item.unitPrice }}</number>
  <formula><rc c="-2" /> * <rc c="-1" /></formula>
</row>
```

## Motivation

Often, you already have a dataset that you have rendered:

```html
<table>
  <thead>
    <tr>
      <th>Initial</th>
      <th>Name</th>
      <th>Birthday</th>
    </tr>
  </thead>
  <tbody>
    <tr v-for="row in dataset">
      <td>{{ row['name'][0] }}</td>
      <td>{{ row['name'] }}</td>
      <td>{{ row['birthday'] }}</td>
    </tr>
  </tbody>
</table>
```

Of course, simple tables don't stay simple for very long. You add sorting:

```html
  <th><SortByButton field="name" :sortField="currentSortField" />Name</th>

  <tr v-for="row in sortBy(dataset, currentSortField)">
```

You add some filtering:

```html
  <tr v-for="row in filterBy(sortBy(dataset, currentSortField), filterCriteria)">
```

You add some rendering:

```html
  <td>
    {{ renderDate(row['birthday'], 'yyyy-mm-dd') }}
    <template v-if="isSameDay(row['birthday'], today)">🎂</template>
  </td>
```

All is well and good, until Business says, **export this to Excel**.

You have two options:
1. **Option 1: Export to Excel in the backend.** However, you quickly reject this option, because you don't want to have to re-implement sorting and filtering in the backend.
2. **Option 2: Export to Excel in the frontend.** This is viable with [SheetJS](https://github.com/SheetJS/sheetjs). So you try it:

## Exporting to Excel with SheetJS
You have three options, in increasing level of complexity:

### (Easy, with major constraints) Table-to-HTML conversion
You use the [`table_to_sheet`](https://docs.sheetjs.com/docs/api/utilities/html#html-table-input) utility method in SheetJS.

```html
<table :ref="refToTable">
```

```js
const refToTable = ref(null);

const convertToSheet = () => XLSX.utils.table_to_sheet(refToTable.value, {})
```

This converts an entire HTML table to a Excel sheet.

Pros:
* Matches your HTML output!

Cons:
* Your HTML must match the intended Excel table exactly. Remember the "cake" that we added to the "birthday" column? It's going to mess up your Excel :(
* You can't specify with precision how you want each field rendered -- e.g. is "6-4" a date or a string?

### (Less easy, with some more constraints) Dataset-to-HTML conversion

So, instead of using something as imprecise as HTML, we use JSON (and allow dates!):

```js
const dataset = ref([
  /* ID and country are fields in the dataset not sent to the presentation layer */
  {id: 10, name: 'Michael', birthday: new Date(1990, 1, 1), country: 'US'},
  {id: 20, name: 'David', birthday: new Date(1990, 6, 5), country: 'SG'},
])

const filteredAndSorted = computed(() => filterBy(sortBy(dataset.value, currentSortField.value), filterCriteria.value))

const convertToSheet = () => XLSX.utils.json_to_sheet(filteredAndSorted.value, opts)
```

Pros:
* Matches the underlying HTML _dataset_

Cons:
* Does not match the HTML output anymore =(. For example, you end up with the extra fields `id` and `country`. In addition, you lose the computed column `initial`.

The Con here is an interesting one -- we can easily work around this by writing a function:
```js
const convertDatasetToExcelDataset = (row) => ({
  initial: row.name[0],
  name: row.name,
  ...
})
```

but that feels awfully similar to the Vue template we wrote earlier, doesn't it? **Why would we want to repeat "presentation" code, once in the template, and once more in Javascript?**

**Question: Can we keep the presentation code inside the template?**

### Proposed solution: Keep presentation code inside template

```html
<DownloadExcelButton>
  <!-- here, we define how we want the spreadsheet presented -->
  <row>
    <text>Initial</text>
    <text>Name</text>
    <text>Birthday</text>
    <text>Age this year</text>
  </row>
  <row v-for="row in dataset">
    <text>{{ row['name'][0] }}</text>
    <text>{{ row['name'] }}</text>
    <date z="MMM DD">{{ row['birthday'] }}</date>
    <number>{{ today.getUTCFullYear() - row['birthday'].getUTCFullYear() }}</number>
  </row>
</DownloadExcelButton>

<!-- here, we define how the page is presented -->
<table>
  <thead>
    <tr>
      <th>Initial</th>
      <th>Name</th>
      <th>Birthday</th>
    </tr>
  </thead>
  <tbody>
    <tr v-for="row in dataset">
      <td>{{ row['name'][0] }}</td>
      <td>{{ row['name'] }}</td>
      <td>{{ row['birthday'] }}</td>
    </tr>
  </tbody>
</table>
```

Pros:
- Clear separation between "presentation code" (in template) vs "processing code" (Javascript)
- Control over formatting details. For example, I can instruct the spreadsheet to render the date as `MMM DD`
- Control over data types -- I can coerce the data types, e.g. with `<text>` and `<number>`, so that the name of a company named "1e1" isn't interpreted as `10`.

Cons:
- We end up duplicating the "presentation code". However, this is still far better than the previous solution, because we keep both the Excel and HTML parts in the "presentation" section with similar syntax, so we can copy and paste code between parts if necessary.

## Q&A

### Why not introduce custom components, instead of introducing new, invalid DOM elements?

**Option 1: parse children**

I could have use custom components like so:

```html
<template>
  <SheetJsOutput>
    <ExcelRow>
      <ExcelTextCell>
```

However, these components would still have to be compiled into a VDOM node. So, what would the template of the `ExcelRow` component be? If it were empty, it'd still result in an empty child VNode. If it weren't empty, it'd either have to be DOM-targeting, or target some other object model.

Could we have tried parsing the children directly, like how [react-export-excel](https://github.com/rdcalle/react-export-excel/blob/5ede0a82492563f9219ba8d53cda406c22811347/src/ExcelPlugin/components/ExcelFile.js#L23-L28) does it? Firstly, Vue 3 doesn't offer access to `$children`. Secondly, even if it did, this is shifts the responsibility of parsing from Vue to the component. If the component dictates a certain structure of elements inside its slot / children, then it prevents users from using custom components to extract repeating elements out of their spreadsheet.

For example, this wouldn't be possible:

```html
<SheetJsOutput>
  <TwoCurrencies :currencies="['usd', 'jpy']" v-for="entry in entries" :data="entry" />
```

where
```html
<!-- TwoCurrencies.vue -->
<ExcelRow v-for="currency in currencies">
  <ExcelNumberCell>{{ convert(data.amount, currency) }}</ExcelNumberCell>
  <ExcelTextCell>{{ currency }}</ExcelTextCell>
</ExcelRow>
```

because SheetJsOutput would be looking for **immediate children** that are instances of ExcelRow, not a custom element like TwoCurrencies.

**Option 2: use lifecycle hooks to add/remove rows**

This is the mechanism by which vue3-google-maps (and vue2-google-maps) inserted elements to the Google Maps layer without resorting to a custom document object model (XDOM).

Unfortunately, this worked only because on a Map, the **order of appearance** of child elements does not matter. That is, it doesn't matter that Marker A is inserted before or after Marker B, as long as both markers appear on the map.
In a spreadsheet though, the order of elements do matter -- you don't want your revenue items appearing after the "Expenses" heading. Worse, it's possible that child elements are re-ordered without any lifecycle hooks being called, e.g. `v-for="r in sortBy(rows, sortField)"`, if we change `sortField`, none of the child elements will be torn down, but the order of appearance should be changed. There is also no information telling us at what index a custom element is being rendered.

Thus, using lifecycle hooks is not a feasible option where sequence of elements matters to the custom object model.

### Do we have to implement XDOM-manipulation?

In `renderer.ts` I implemented lots of custom document object model (XDOM) manipulation methods like `patchProp`, `insert`, `remove` etc. which isn't strictly necessary.

We could really just parse VNodes directly, e.g. `vnode.children.map(x => { if (x.type === 'row') { ... } })`. However implementing an XDOM allows for a bit nicer type readability and safety (observe all the types defined in the same file). With a little bit more effort, I could even define the known attributes like `width` and `z` on the types themselves.

On the other hand if we had parsed VNodes directly, all the types and attributes would have been implicit in the parsing method, resulting in code that's harder to read and understand.
