import { mount } from '@vue/test-utils'
import { nextTick } from 'vue'
import SheetRenderer from './TestSheetRenderer.vue'

test('renders the expected Sheet object', async () => {
  const wrapper = mount(SheetRenderer)

  await nextTick()

  const text = wrapper.text()

  expect(JSON.parse(text)).toEqual({
    "!ref": "A1:H3",
    "!cols": [
      {},
      {},
      { wch: 11 },
      {},
      {},
      {},
      { wch: 20 },
      { wch: 20 },
    ],
    "A1": { "t": "s", "v": "Employee ID" },
    "A2": { "t": "s", "v": "101" },
    "A3": { "t": "s", "v": "102" },
    "B1": { "t": "s", "v": "Name" },
    "B2": { "t": "s", "v": "Matthew" },
    "B3": { "t": "s", "v": "Julius" },
    "C1": { "t": "s", "v": "Date Joined" },
    "C2": { "t": "n", "v": 44713, "z": "YYYY-MM-DD" },
    "C3": { "t": "n", "v": 45356, "z": "YYYY-MM-DD" },
    "D1": { "t": "s", "v": "YoE at Hire" },
    "D2": { "t": "n", "v": 5.5, "z": "#,##0.0" },
    "D3": { "t": "n", "v": 3, "z": "#,##0.0" },
    "E1": { "t": "s", "v": "YoE to date" },
    "E2": { "f": "D2 + (now() - C2) / 365 " },
    "E3": { "f": "D3 + (now() - C3) / 365 " },
    "F1": { "t": "s", "v": "Is Manager?" },
    "F2": { "t": "b", "v": false },
    "F3": { "t": "b", "v": true },
    "G1": { "t": "s", "v": "Created" },
    "G2": { "t": "n", "v": 44711.67303240741, "z": "YYYY-MM-DD" },
    "G3": { "t": "n", "v": 45352.29803240741, "z": "YYYY-MM-DD" },
    "H1": { "t": "s", "v": "Updated" },
    "H2": { "t": "n", "v": 44711.67303240741, "z": "YYYY-MM-DD" },
    "H3": { "t": "n", "v": 45352.29803240741, "z": "YYYY-MM-DD" }
  })
})