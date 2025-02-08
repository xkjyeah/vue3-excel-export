// This renderer simply stores the VNodes, which we can then parse out
// into a AoA + AoF that we can pass to SheetJS

import { createRenderer } from '@vue/runtime-core'
import type { VNodeProps, ElementNamespace, ComponentInternalInstance } from '@vue/runtime-core'
import { utils as XLSXutils } from "xlsx";

type BaseNode = {
  previousSibling: XlNode | null,
  nextSibling: XlNode | null,
  parentElement: XlElement | null,
}

class BaseElement implements BaseNode {
  attributes: (VNodeProps & { [key: string]: any }) = {};
  children: XlNode[] = [];
  namespace?: string;
  previousSibling: XlNode | null = null;
  nextSibling: XlNode | null = null;
  parentElement: XlElement | null = null;
  type: string;

  constructor(type: string) {
    this.attributes = {}
    this.children = []
    this.type = type
  }

  getText() {
    return this.children.filter((x): x is TextNode => !(x instanceof BaseElement) && x.type === 'text-node')
      .map(tn => tn.data.toString())
      .join('')
  }
}

class BaseCell extends BaseElement { }

export class TextCell extends BaseCell {
  constructor() {
    super('text')
  }

  data(): any {
    return {
      t: 's',
      v: this.getText()
    }
  }
}
export class BooleanCell extends BaseCell {
  constructor() {
    super('boolean')
  }

  data(): any {
    const t = this.getText()

    return (t !== '') ? {
      t: 'b',
      v: t.trim() === 'true'
    } : null
  }
}
export class NumberCell extends BaseCell {
  constructor() {
    super('number')
  }

  data(): any {
    const t = this.getText()

    return (t !== '') ? {
      t: 'n',
      v: parseFloat(t.trim())
    } : null
  }
}
export class DateCell extends BaseCell {
  constructor() {
    super('date')
  }


  data(): any {
    const t = this.getText()

    if (t === '') {
      return null
    }

    const d = new Date(t)
    const basedate = new Date('1899-12-31T00:00:00Z');

    const adjustmentBasedate = new Date('1900-03-01T00:00:00Z');
    const adjustmentFor1900LeapYearBug = d.getTime() > adjustmentBasedate.getTime() ? 86400e3 : 0

    return {
      t: 'n',
      v: (d.getTime() - basedate.getTime() + adjustmentFor1900LeapYearBug) /
        86400e3,
      z: 'YYYY-MM-DD'
    }
  }
}
export class RowElement extends BaseElement {
  constructor() {
    super('row')
  }
}
export class SheetElement extends BaseElement {
  constructor() {
    super('sheet')
  }
}
export class FormulaCell extends BaseElement {
  constructor() {
    super('formula')
  }
}
export type CommentNode = BaseNode & {
  type: 'comment',
  data: string
}
export type TextNode = BaseNode & {
  type: 'text-node',
  data: string
}
export type XlElement = TextCell | BooleanCell | NumberCell | DateCell | FormulaCell | RowElement | SheetElement
export type XlNode = XlElement | TextNode | CommentNode


export const EMPTY_NODE_RELATIONSHIP_DATA: BaseNode = {
  parentElement: null,
  previousSibling: null,
  nextSibling: null,
}
export const EMPTY_ELEMENT_RELATIONSHIP_DATA: BaseNode & Pick<BaseElement, 'children'> = {
  ...EMPTY_NODE_RELATIONSHIP_DATA,
  children: [],
}

const isPresent = function <T extends Object>(t: T | undefined | null): t is T {
  return (t !== undefined && t !== null)
}

const toData = (f: XlNode) => {
  const c = (f instanceof BaseElement) ? f.children : []

  const baseSheet = XLSXutils.aoa_to_sheet(c.map((row: XlNode, index: number) => {
    if (!(row instanceof RowElement)) {
      console.warn(`Unclear how to handle child of sheet. Skipping...`, row.type)
      return
    }

    return row.children.map((maybeCellElem: XlNode) => {
      if (maybeCellElem instanceof TextCell ||
        maybeCellElem instanceof BooleanCell ||
        maybeCellElem instanceof DateCell ||
        maybeCellElem instanceof NumberCell
      ) {
        const format = maybeCellElem.attributes['z']
        const result = maybeCellElem.data()

        if (format && result) {
          result.z = format
        }
        return result
      } else if (maybeCellElem instanceof FormulaCell) {
        // Mark it as blank for now...
        // TODO: return {t: '??', f: '<formula>'} Cell object
        // Note: it would be nice to allow it to reference cells in R1C1 format, but
        // that's not supported by SheetJS now??
        return null
      }
    })
  }).filter(isPresent)
  )

  // Any column objects
  const widthSettingRow = c.find((r: XlNode): r is RowElement =>
    r instanceof RowElement && r.attributes['widthSetting'])

  const cols = widthSettingRow?.children?.filter((r: XlNode): r is BaseCell =>
    r instanceof BaseCell)
    .map(r => {
      const { width } = r.attributes
      const result: any = {}

      if (width) {
        result.wch = parseInt(width)
      }

      return result
    })

  if (cols) {
    baseSheet['!cols'] = cols
  }

  return baseSheet
}

const { render, createApp } = createRenderer<XlNode, XlElement>({
  /*
  A bunch of boilerplate to track the VDOM result. The VDOM
  result will be converted directly to a spreadsheet.

  It's possible that none of this is necessary, and that we parse the VDOM
  directly
  */
  patchProp(
    el: XlElement,
    key: string,
    prevValue: any,
    nextValue: any,
    namespace?: ElementNamespace,
    parentComponent?: ComponentInternalInstance | null,
  ) {
    el.attributes[key] = nextValue
  },
  insert(el: XlNode, parent: XlElement, insertBefore?: XlNode | null) {
    let indexOfAnchor = insertBefore ? parent.children.indexOf(insertBefore) : parent.children.length

    if (indexOfAnchor === -1) {
      throw new Error('Trying to insert before child node that does not exist')
    }

    // Update sibling relationship
    const prev = parent.children[indexOfAnchor - 1]
    const next = parent.children[indexOfAnchor]

    if (prev) {
      prev.nextSibling = el
    }
    if (next) {
      next.previousSibling = el
    }
    el.previousSibling = prev
    el.nextSibling = next

    // Update parent relationship
    parent.children.splice(indexOfAnchor, 0, el)
    el.parentElement = parent
  },
  remove(el: XlNode) {
    // Disconnect from parent.
    const parent = el.parentElement

    if (parent) {
      parent.children = parent.children.filter(c => c !== el)
    }

    // Disconnect from siblings
    const prev = el.previousSibling
    const next = el.nextSibling

    if (prev) {
      prev.nextSibling = next
    }
    if (next) {
      next.previousSibling = prev
    }
    Object.assign(el, EMPTY_NODE_RELATIONSHIP_DATA)
  },
  createElement(
    type: XlElement['type'],
    namespace?: ElementNamespace,
    isCustomizedBuiltIn?: string,
    vnodeProps?: (VNodeProps & { [key: string]: any }) | null,
  ): XlElement {
    switch (type) {
      case 'text':
        return new TextCell()
      case 'row':
        return new RowElement()
      case 'boolean':
        return new BooleanCell()
      case 'date':
        return new DateCell()
      case 'number':
        return new NumberCell()
      case 'formula':
        return new FormulaCell()
      default:
        throw new Error(`Unsupported element type ${type}`)
    }
  },
  createText(text: string): TextNode {
    return {
      type: 'text-node',
      data: text,
      ...EMPTY_NODE_RELATIONSHIP_DATA
    }
  },
  createComment(text: string): CommentNode {
    debugger
    return {
      type: 'comment',
      data: text,
      ...EMPTY_NODE_RELATIONSHIP_DATA
    }
  },
  setText(node: XlNode, text: string) {
    // There's nothing much to do here -- since we're
    // not updating any canvas, we just need to store the data
    if (node.type !== 'text-node' || node instanceof BaseElement) {
      throw new Error('Expected text node!')
    }
    node.data = text
  },
  setElementText(node: XlElement, text: string) {
    // Disconnect all internal nodes
    node.children.forEach(f => {
      Object.assign(f, EMPTY_NODE_RELATIONSHIP_DATA)
    })

    // Update the text
    node.children = [
      {
        type: 'text-node',
        data: text,
        ...EMPTY_NODE_RELATIONSHIP_DATA,
      }
    ]
  },
  parentNode(node: XlNode) {
    return node.parentElement
  },
  nextSibling(node: XlNode) {
    return node.nextSibling
  },
})


export { render, createApp, toData }
