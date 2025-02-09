import { createApp, toData } from "./renderer"
import { SheetElement } from "./renderer"
import type { Component } from 'vue'
import { ref, onMounted } from 'vue'

export default <Component>{
  setup(props, ctx) {
    const docElement = ref<SheetElement>(new SheetElement())

    onMounted(() => {
      const sheetSlot = ctx.slots.sheet

      createApp({
        setup() {
          return () => sheetSlot && sheetSlot()
        }
      })
        .mount(docElement.value)
    })

    return () => {
      const defaultSlot = ctx.slots.default
      return defaultSlot ? defaultSlot(toData(docElement.value)) : null
    }
  },
}
