# vue3-excel-export

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

