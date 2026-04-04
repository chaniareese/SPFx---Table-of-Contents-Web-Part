import * as React from 'react';
import { useEditor, EditorContent, ReactNodeViewRenderer, NodeViewWrapper } from '@tiptap/react';
import type { ReactNodeViewProps } from '@tiptap/react';
import { Editor, Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import StarterKit from '@tiptap/starter-kit';
import { Table } from '@tiptap/extension-table';
import TableRow from '@tiptap/extension-table-row';
import TableCell from '@tiptap/extension-table-cell';
import TableHeader from '@tiptap/extension-table-header';
import Underline from '@tiptap/extension-underline';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyle } from '@tiptap/extension-text-style';
import Color from '@tiptap/extension-color';
import Highlight from '@tiptap/extension-highlight';
import Link from '@tiptap/extension-link';
import Image from '@tiptap/extension-image';
import Superscript from '@tiptap/extension-superscript';
import Subscript from '@tiptap/extension-subscript';
import Placeholder from '@tiptap/extension-placeholder';
import styles from './TipTapEditor.module.scss';

/* ------------------------------------------------------------------ */
/*  Module Augmentation – declare custom commands for TypeScript       */
/* ------------------------------------------------------------------ */
declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    fontSize: {
      setFontSize: (size: string) => ReturnType;
      unsetFontSize: () => ReturnType;
    };
    lineHeight: {
      setLineHeight: (height: string) => ReturnType;
      unsetLineHeight: () => ReturnType;
    };
    indent: {
      increaseIndent: () => ReturnType;
      decreaseIndent: () => ReturnType;
    };
    bulkCellAttribute: {
      setBulkCellAttribute: (attr: string, value: unknown) => ReturnType;
    };
  }
}

/* ------------------------------------------------------------------ */
/*  Custom Extensions                                                  */
/* ------------------------------------------------------------------ */

const FontSize = Extension.create({
  name: 'fontSize',
  addOptions() {
    return { types: ['textStyle'] };
  },
  addGlobalAttributes() {
    return [{
      types: this.options.types,
      attributes: {
        fontSize: {
          default: null,
          parseHTML: (element: HTMLElement) => element.style.fontSize || null,
          renderHTML: (attributes: Record<string, unknown>) => {
            if (!attributes.fontSize) return {};
            return { style: `font-size: ${attributes.fontSize}` };
          },
        },
      },
    }];
  },
  addCommands() {
    return {
      setFontSize: (fontSize: string) => ({ chain }) => {
        return chain().setMark('textStyle', { fontSize }).run();
      },
      unsetFontSize: () => ({ chain }) => {
        return chain().setMark('textStyle', { fontSize: null }).removeEmptyTextStyle().run();
      },
    };
  },
});

const LineHeight = Extension.create({
  name: 'lineHeight',
  addOptions() {
    return { types: ['paragraph', 'heading'] };
  },
  addGlobalAttributes() {
    return [{
      types: this.options.types,
      attributes: {
        lineHeight: {
          default: null,
          parseHTML: (element: HTMLElement) => element.style.lineHeight || null,
          renderHTML: (attributes: Record<string, unknown>) => {
            if (!attributes.lineHeight) return {};
            return { style: `line-height: ${attributes.lineHeight}` };
          },
        },
      },
    }];
  },
  addCommands() {
    return {
      setLineHeight: (lineHeight: string) =>
        ({ tr, state, dispatch }) => {
          const { from, to } = state.selection;
          let updated = false;
          state.doc.nodesBetween(from, to, (node, pos) => {
            if (this.options.types.includes(node.type.name)) {
              tr.setNodeMarkup(pos, undefined, { ...node.attrs, lineHeight });
              updated = true;
            }
          });
          if (updated && dispatch) dispatch(tr);
          return updated;
        },
      unsetLineHeight: () =>
        ({ tr, state, dispatch }) => {
          const { from, to } = state.selection;
          let updated = false;
          state.doc.nodesBetween(from, to, (node, pos) => {
            if (this.options.types.includes(node.type.name) && node.attrs.lineHeight) {
              tr.setNodeMarkup(pos, undefined, { ...node.attrs, lineHeight: null });
              updated = true;
            }
          });
          if (updated && dispatch) dispatch(tr);
          return updated;
        },
    };
  },
});

const Indent = Extension.create({
  name: 'indent',
  addOptions() {
    return { types: ['paragraph', 'heading'], listTypes: ['listItem'], min: 0, max: 6 };
  },
  addGlobalAttributes() {
    return [
      {
        // paragraph and heading — use margin-left
        types: this.options.types,
        attributes: {
          indent: {
            default: 0,
            parseHTML: (element: HTMLElement) => {
              const ml = parseInt(element.style.marginLeft || '0', 10);
              return Math.round(ml / 40) || 0;
            },
            renderHTML: (attributes: Record<string, unknown>) => {
              const indent = attributes.indent as number;
              if (!indent || indent <= 0) return {};
              return { style: `margin-left: ${indent * 40}px` };
            },
          },
        },
      },
      {
        // listItem — use padding-left so the marker moves WITH the text
        types: this.options.listTypes,
        attributes: {
          indent: {
            default: 0,
            parseHTML: (element: HTMLElement) => {
              const pl = parseInt(element.style.paddingLeft || '0', 10);
              return Math.round(pl / 40) || 0;
            },
            renderHTML: (attributes: Record<string, unknown>) => {
              const indent = attributes.indent as number;
              if (!indent || indent <= 0) return {};
              return { style: `padding-left: ${indent * 40}px` };
            },
          },
        },
      },
    ];
  },
  addCommands() {
    const allTypes = [...this.options.types, ...this.options.listTypes];
    return {
      increaseIndent: () =>
        ({ tr, state, dispatch }) => {
          const { from, to } = state.selection;
          let updated = false;
          state.doc.nodesBetween(from, to, (node, pos) => {
            if (allTypes.includes(node.type.name)) {
              const cur = (node.attrs.indent as number) || 0;
              if (cur < this.options.max) {
                tr.setNodeMarkup(pos, undefined, { ...node.attrs, indent: cur + 1 });
                updated = true;
              }
            }
          });
          if (updated && dispatch) dispatch(tr);
          return updated;
        },
      decreaseIndent: () =>
        ({ tr, state, dispatch }) => {
          const { from, to } = state.selection;
          let updated = false;
          state.doc.nodesBetween(from, to, (node, pos) => {
            if (allTypes.includes(node.type.name)) {
              const cur = (node.attrs.indent as number) || 0;
              if (cur > this.options.min) {
                tr.setNodeMarkup(pos, undefined, { ...node.attrs, indent: cur - 1 });
                updated = true;
              }
            }
          });
          if (updated && dispatch) dispatch(tr);
          return updated;
        },
    };
  },
});

/* ------------------------------------------------------------------ */
/*  Extended Table nodes – custom cell attributes                      */
/* ------------------------------------------------------------------ */

const CustomTableCell = TableCell.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      backgroundColor: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.backgroundColor || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.backgroundColor ? { style: `background-color: ${a.backgroundColor}` } : {},
      },
      borderStyle: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.borderStyle || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.borderStyle ? { style: `border-style: ${a.borderStyle}` } : {},
      },
      borderColor: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.borderColor || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.borderColor ? { style: `border-color: ${a.borderColor}` } : {},
      },
      cellPadding: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.padding || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.cellPadding ? { style: `padding: ${a.cellPadding}` } : {},
      },
      verticalAlign: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.verticalAlign || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.verticalAlign ? { style: `vertical-align: ${a.verticalAlign}` } : {},
      },
    };
  },
});

const CustomTableHeader = TableHeader.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      backgroundColor: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.backgroundColor || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.backgroundColor ? { style: `background-color: ${a.backgroundColor}` } : {},
      },
      borderStyle: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.borderStyle || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.borderStyle ? { style: `border-style: ${a.borderStyle}` } : {},
      },
      borderColor: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.borderColor || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.borderColor ? { style: `border-color: ${a.borderColor}` } : {},
      },
      cellPadding: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.padding || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.cellPadding ? { style: `padding: ${a.cellPadding}` } : {},
      },
      verticalAlign: {
        default: null,
        parseHTML: (el: HTMLElement) => el.style.verticalAlign || null,
        renderHTML: (a: Record<string, unknown>) =>
          a.verticalAlign ? { style: `vertical-align: ${a.verticalAlign}` } : {},
      },
    };
  },
});

/* ------------------------------------------------------------------ */
/*  Bulk Cell Attribute Extension                                      */
/*  Applies a cell attribute to ALL selected cells, not just current  */
/*  Handles both CellSelection (multi-cell drag) and cursor position  */
/* ------------------------------------------------------------------ */

const BulkCellAttribute = Extension.create({
  name: 'bulkCellAttribute',
  addCommands() {
    return {
      setBulkCellAttribute: (attr: string, value: unknown) =>
        ({ tr, state, dispatch }) => {
          const { selection } = state;
          const cellNodeTypes = ['tableCell', 'tableHeader'];
          let updated = false;

          // Use constructor name — $cellAnchor is on prototype, not own property
          const isCellSelection = selection.constructor?.name === 'CellSelection';

          if (isCellSelection) {
            // Iterate from selection.from to selection.to
            // CellSelection exposes standard from/to that covers all selected cells
            state.doc.nodesBetween(selection.from, selection.to, (node, pos) => {
              if (cellNodeTypes.includes(node.type.name)) {
                tr.setNodeMarkup(pos, undefined, { ...node.attrs, [attr]: value });
                updated = true;
                return false; // don't descend into cell content
              }
              return true;
            });
          }

          // Fallback for single cell / regular cursor
          if (!updated) {
            const { $from } = selection;
            for (let d = $from.depth; d >= 0; d--) {
              const node = $from.node(d);
              if (cellNodeTypes.includes(node.type.name)) {
                const pos = $from.before(d);
                tr.setNodeMarkup(pos, undefined, { ...node.attrs, [attr]: value });
                updated = true;
                break;
              }
            }
          }

          if (updated && dispatch) dispatch(tr);
          return updated;
        },
    };
  },
});

/* ------------------------------------------------------------------ */
/*  Cell Position Tracker                                              */
/*  ProseMirror plugin — tracks selected cell positions INSIDE the    */
/*  editor before focus is lost. Stored in module-level array so      */
/*  TableToolbar can read them after blur.                            */
/* ------------------------------------------------------------------ */

const lastSelectedCellPositions: number[] = [];
const cellTrackerKey = new PluginKey('cellTracker');

const CellTracker = Extension.create({
  name: 'cellTracker',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: cellTrackerKey,
        view() {
          return {
            update(view) {
              const { selection } = view.state;
              const cellNodeTypes = ['tableCell', 'tableHeader'];
              const isCellSel = selection.constructor?.name === 'CellSelection';
              if (isCellSel) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const ranges = (selection as any).ranges as Array<{ $from: any }>;
                lastSelectedCellPositions.length = 0;
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                ranges.forEach((range: any) => {
                  const $from = range.$from;
                  for (let d = $from.depth; d >= 0; d--) {
                    if (cellNodeTypes.includes($from.node(d).type.name)) {
                      const pos = $from.before(d);
                      if (!lastSelectedCellPositions.includes(pos)) {
                        lastSelectedCellPositions.push(pos);
                      }
                      break;
                    }
                  }
                });
              }
            }
          };
        }
      })
    ];
  }
});

/* ------------------------------------------------------------------ */
/*  Constants                                                          */
/* ------------------------------------------------------------------ */

const COLORS = [
  '#000000','#434343','#666666','#999999','#b7b7b7','#cccccc','#d9d9d9','#efefef','#f3f3f3','#ffffff',
  '#980000','#ff0000','#ff9900','#ffff00','#00ff00','#00ffff','#4a86e8','#0000ff','#9900ff','#ff00ff',
  '#e6b8af','#f4cccc','#fce5cd','#fff2cc','#d9ead3','#d0e0e3','#c9daf8','#cfe2f3','#d9d2e9','#ead1dc',
  '#dd7e6b','#ea9999','#f9cb9c','#ffe599','#b6d7a8','#a2c4c9','#a4c2f4','#9fc5e8','#b4a7d6','#d5a6bd',
  '#cc4125','#e06666','#f6b26b','#ffd966','#93c47d','#76a5af','#6d9eeb','#6fa8dc','#8e7cc3','#c27ba0',
  '#a61c00','#cc0000','#e69138','#f1c232','#6aa84f','#45818e','#3c78d8','#3d85c6','#674ea7','#a64d79',
];

const FONT_SIZE_OPTIONS: number[] = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 42, 68];

const LINE_HEIGHTS: { label: string; value: string }[] = [
  { label: '1.0', value: '1' },
  { label: '1.5', value: '1.5' },
  { label: '2.0', value: '2' },
];

const BORDER_STYLES: { label: string; value: string }[] = [
  { label: 'Solid', value: 'solid' },
  { label: 'Dashed', value: 'dashed' },
  { label: 'None', value: 'none' },
];

const CELL_PADDINGS: { label: string; value: string }[] = [
  { label: 'Compact', value: '4px 6px' },
  { label: 'Normal', value: '8px 12px' },
  { label: 'Relaxed', value: '12px 16px' },
  { label: 'Spacious', value: '16px 24px' },
];

/* ------------------------------------------------------------------ */
/*  Inline SVG Icons                                                   */
/* ------------------------------------------------------------------ */

const I = {
  alignLeft: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
      <path d="M2 3h12M2 6.3h8M2 9.6h10M2 13h6"/>
    </svg>
  ),
  alignCenter: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
      <path d="M2 3h12M4 6.3h8M3 9.6h10M5 13h6"/>
    </svg>
  ),
  alignRight: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
      <path d="M2 3h12M6 6.3h8M4 9.6h10M8 13h6"/>
    </svg>
  ),
  alignJustify: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
      <path d="M2 3h12M2 6.3h12M2 9.6h12M2 13h12"/>
    </svg>
  ),
  bulletList: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
      <circle cx="3" cy="4" r="1.3"/><rect x="6" y="3" width="8" height="1.6" rx=".5"/>
      <circle cx="3" cy="8" r="1.3"/><rect x="6" y="7" width="8" height="1.6" rx=".5"/>
      <circle cx="3" cy="12" r="1.3"/><rect x="6" y="11" width="8" height="1.6" rx=".5"/>
    </svg>
  ),
  orderedList: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
      <text x="1" y="5.5" fontSize="6" fontFamily="sans-serif">1.</text><rect x="6" y="3" width="8" height="1.6" rx=".5"/>
      <text x="1" y="9.5" fontSize="6" fontFamily="sans-serif">2.</text><rect x="6" y="7" width="8" height="1.6" rx=".5"/>
      <text x="1" y="13.5" fontSize="6" fontFamily="sans-serif">3.</text><rect x="6" y="11" width="8" height="1.6" rx=".5"/>
    </svg>
  ),
  link: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round">
      <path d="M6.8 9.2l2.4-2.4M5 8l-1.6 1.6a2.1 2.1 0 003 3L8 11M11 8l1.6-1.6a2.1 2.1 0 00-3-3L8 5"/>
    </svg>
  ),
  image: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="3" width="12" height="10" rx="1"/><circle cx="5.5" cy="6.5" r="1.3"/>
      <path d="M2 11l3-3 2 2 3-3 4 4" strokeLinejoin="round"/>
    </svg>
  ),
  table: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="2" width="12" height="12" rx="1"/>
      <line x1="2" y1="6" x2="14" y2="6"/><line x1="2" y1="10" x2="14" y2="10"/>
      <line x1="6" y1="2" x2="6" y2="14"/><line x1="10" y1="2" x2="10" y2="14"/>
    </svg>
  ),
  indent: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round">
      <path d="M8 3h6M8 6.5h6M8 10h6M8 13.5h6M2 6.5l3 2-3 2"/>
    </svg>
  ),
  outdent: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round">
      <path d="M8 3h6M8 6.5h6M8 10h6M8 13.5h6M5 6.5l-3 2 3 2"/>
    </svg>
  ),
  clearFormat: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
      <path d="M2 3h8M6 3v10"/>
      <path d="M10 8l4 4M14 8l-4 4" strokeWidth="1.4"/>
    </svg>
  ),
  highlight: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round">
      <path d="M3.5 10.5L9.5 4.5l2 2-6 6H3.5v-2z"/>
      <path d="M9.5 4.5l1.5-1.5 2 2-1.5 1.5"/>
    </svg>
  ),
  // Table icons
  addRowAbove: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="6" width="12" height="8" rx="1"/><line x1="2" y1="10" x2="14" y2="10"/><line x1="8" y1="6" x2="8" y2="14"/>
      <path d="M8 1v3M6.5 2.5h3" strokeWidth="1.5" strokeLinecap="round"/>
    </svg>
  ),
  addRowBelow: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="2" width="12" height="8" rx="1"/><line x1="2" y1="6" x2="14" y2="6"/><line x1="8" y1="2" x2="8" y2="10"/>
      <path d="M8 12v3M6.5 13.5h3" strokeWidth="1.5" strokeLinecap="round"/>
    </svg>
  ),
  addColLeft: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="6" y="2" width="8" height="12" rx="1"/><line x1="6" y1="8" x2="14" y2="8"/><line x1="10" y1="2" x2="10" y2="14"/>
      <path d="M2.5 8h-2M1.5 6.5v3" strokeWidth="1.5" strokeLinecap="round"/>
    </svg>
  ),
  addColRight: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="2" width="8" height="12" rx="1"/><line x1="2" y1="8" x2="10" y2="8"/><line x1="6" y1="2" x2="6" y2="14"/>
      <path d="M13.5 8h2M14.5 6.5v3" strokeWidth="1.5" strokeLinecap="round"/>
    </svg>
  ),
  deleteRow: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="2" width="12" height="12" rx="1"/><line x1="2" y1="6" x2="14" y2="6"/><line x1="2" y1="10" x2="14" y2="10"/>
      <path d="M5 7.2l6 1.6M5 8.8l6-1.6" stroke="#a4262c" strokeWidth="1.4" strokeLinecap="round"/>
    </svg>
  ),
  deleteCol: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="2" width="12" height="12" rx="1"/><line x1="6" y1="2" x2="6" y2="14"/><line x1="10" y1="2" x2="10" y2="14"/>
      <path d="M7.2 5l1.6 6M8.8 5l-1.6 6" stroke="#a4262c" strokeWidth="1.4" strokeLinecap="round"/>
    </svg>
  ),
  deleteTable: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="#a4262c" strokeWidth="1.2">
      <rect x="2" y="2" width="12" height="12" rx="1"/>
      <line x1="2" y1="6" x2="14" y2="6"/><line x1="2" y1="10" x2="14" y2="10"/>
      <line x1="6" y1="2" x2="6" y2="14"/><line x1="10" y1="2" x2="10" y2="14"/>
      <path d="M4 4l8 8M12 4l-8 8" strokeWidth="1.6" strokeLinecap="round"/>
    </svg>
  ),
  mergeCells: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="2" width="12" height="12" rx="1"/>
      <path d="M5 8h6M9.5 5.5L12 8l-2.5 2.5M6.5 5.5L4 8l2.5 2.5" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.3"/>
    </svg>
  ),
  splitCell: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.2">
      <rect x="2" y="2" width="12" height="12" rx="1"/>
      <path d="M4 8h3M9 8h3M4.5 5.5L2 8l2.5 2.5M11.5 5.5L14 8l-2.5 2.5" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.3"/>
    </svg>
  ),
  paragraphBefore: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round">
      <path d="M4 7h8M4 10h8M4 13h5"/><path d="M8 1v3M6.5 2L8 0.5 9.5 2"/>
    </svg>
  ),
  paragraphAfter: (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round">
      <path d="M4 3h8M4 6h8M4 9h5"/><path d="M8 12v3M6.5 14L8 15.5 9.5 14"/>
    </svg>
  ),
};

/* ------------------------------------------------------------------ */
/*  Sub‑Components                                                     */
/* ------------------------------------------------------------------ */

interface IColorPickerProps {
  colors: string[];
  activeColor?: string;
  onSelect: (c: string) => void;
  onClear?: () => void;
  onClose: () => void;
}

const ColorPicker: React.FC<IColorPickerProps> = ({ colors, activeColor, onSelect, onClear, onClose }) => {
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [onClose]);

  return (
    <div className={styles.colorPicker} ref={ref}>
      <div className={styles.colorGrid}>
        {colors.map(c => (
          <button
            key={c}
            className={`${styles.colorSwatch} ${activeColor === c ? styles.activeSwatch : ''}`}
            style={{ backgroundColor: c, border: c === '#ffffff' ? '1px solid #ccc' : undefined }}
            onMouseDown={e => e.preventDefault()}
            onClick={() => { onSelect(c); onClose(); }}
            title={c}
          />
        ))}
      </div>
      {onClear && (
        <button className={styles.colorClear} onMouseDown={e => e.preventDefault()} onClick={() => { if (onClear) onClear(); onClose(); }}>
          No color
        </button>
      )}
    </div>
  );
};

interface ITableGridProps { onSelect: (r: number, c: number) => void; onClose: () => void }

const TableGridSelector: React.FC<ITableGridProps> = ({ onSelect, onClose }) => {
  const [hr, setHR] = React.useState(0);
  const [hc, setHC] = React.useState(0);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [onClose]);

  return (
    <div className={styles.gridSelector} ref={ref}>
      <div className={styles.gridLabel}>{hr > 0 ? `${hr} × ${hc}` : 'Insert Table'}</div>
      <div className={styles.grid}>
        {Array.from({ length: 6 }, (_, r) => (
          <div key={r} className={styles.gridRow}>
            {Array.from({ length: 6 }, (_, c) => (
              <div
                key={c}
                className={`${styles.gridCell} ${r < hr && c < hc ? styles.gridHighlight : ''}`}
                onMouseEnter={() => { setHR(r + 1); setHC(c + 1); }}
                onClick={() => { onSelect(r + 1, c + 1); onClose(); }}
              />
            ))}
          </div>
        ))}
      </div>
    </div>
  );
};

interface ILinkInputProps {
  initialUrl?: string;
  onSubmit: (url: string) => void;
  onRemove: () => void;
  onClose: () => void;
}

const LinkInput: React.FC<ILinkInputProps> = ({ initialUrl, onSubmit, onRemove, onClose }) => {
  const [url, setUrl] = React.useState(initialUrl || '');
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [onClose]);

  return (
    <div className={styles.linkInput} ref={ref}>
      <input
        type="url"
        value={url}
        onChange={e => setUrl(e.target.value)}
        placeholder="https://..."
        onKeyDown={e => { if (e.key === 'Enter' && url) { onSubmit(url); onClose(); } }}
        ref={inputRef => { if (inputRef) inputRef.focus(); }}
      />
      <button disabled={!url} onClick={() => { onSubmit(url); onClose(); }}>Apply</button>
      {initialUrl && <button onClick={() => { onRemove(); onClose(); }}>Remove</button>}
    </div>
  );
};

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */

function findTablePos(editor: Editor): number | null {
  const { $from } = editor.state.selection;
  for (let d = $from.depth; d >= 0; d--) {
    if ($from.node(d).type.name === 'table') return $from.before(d);
  }
  return null;
}

function addLineBefore(editor: Editor): void {
  const pos = findTablePos(editor);
  if (pos !== null) {
    editor.chain().focus().insertContentAt(pos, { type: 'paragraph' }).run();
  }
}

function addLineAfter(editor: Editor): void {
  const { $from } = editor.state.selection;
  for (let d = $from.depth; d >= 0; d--) {
    if ($from.node(d).type.name === 'table') {
      const end = $from.before(d) + $from.node(d).nodeSize;
      editor.chain().focus().insertContentAt(end, { type: 'paragraph' }).run();
      return;
    }
  }
}

/* ------------------------------------------------------------------ */
/*  Toolbar button helpers                                             */
/* ------------------------------------------------------------------ */

interface IBtnProps {
  active?: boolean;
  disabled?: boolean;
  title: string;
  onClick: () => void;
  children: React.ReactNode;
  className?: string;
}

const Btn: React.FC<IBtnProps> = ({ active, disabled, title, onClick, children, className }) => (
  <button
    className={`${styles.tbBtn} ${active ? styles.tbActive : ''} ${className || ''}`}
    disabled={disabled}
    title={title}
    onMouseDown={e => e.preventDefault()}
    onClick={onClick}
  >
    {children}
  </button>
);

/* ------------------------------------------------------------------ */
/*  Resizable Image                                                    */
/* ------------------------------------------------------------------ */

interface IResizableImageViewProps extends ReactNodeViewProps {
  // extends ReactNodeViewProps for proper typing with ReactNodeViewRenderer
}

const ResizableImageView: React.FC<IResizableImageViewProps> = ({ node, updateAttributes, selected }) => {
  const imgRef = React.useRef<HTMLImageElement>(null);

  const onResizeStart = (e: React.MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = imgRef.current?.offsetWidth || 200;

    const onMouseMove = (ev: MouseEvent): void => {
      const newWidth = Math.max(50, startWidth + (ev.clientX - startX));
      if (imgRef.current) imgRef.current.style.width = `${newWidth}px`;
    };

    const onMouseUp = (ev: MouseEvent): void => {
      const finalWidth = Math.max(50, startWidth + (ev.clientX - startX));
      updateAttributes({ width: finalWidth });
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  };

  return (
    <NodeViewWrapper className={styles.resizableImageWrapper}>
      <div className={`${styles.resizableImageContainer} ${selected ? styles.resizableImageSelected : ''}`}>
        <img
          ref={imgRef}
          src={node.attrs.src}
          alt={node.attrs.alt || ''}
          title={node.attrs.title || undefined}
          style={node.attrs.width ? { width: `${node.attrs.width}px` } : undefined}
          draggable={false}
        />
        {selected && (
          <div className={styles.resizeHandle} onMouseDown={onResizeStart} />
        )}
      </div>
    </NodeViewWrapper>
  );
};

const ResizableImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      width: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const w = element.getAttribute('width') || element.style.width;
          return w ? parseInt(String(w), 10) : null;
        },
        renderHTML: (attributes: Record<string, unknown>) => {
          if (!attributes.width) return {};
          return { width: String(attributes.width) };
        },
      },
    };
  },
  addNodeView() {
    return ReactNodeViewRenderer(ResizableImageView);
  },
});

/* ------------------------------------------------------------------ */
/*  Image Insert (file upload, paste, or URL)                          */
/* ------------------------------------------------------------------ */

const ImageInsert: React.FC<{ editor: Editor }> = ({ editor }) => {
  const [showPopup, setShowPopup] = React.useState(false);
  const [urlValue, setUrlValue] = React.useState('');
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const popupRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (popupRef.current && !popupRef.current.contains(e.target as Node)) setShowPopup(false);
    };
    if (showPopup) document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [showPopup]);

  const insertDataUrl = (file: File): void => {
    if (!file.type.startsWith('image/')) return;
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      editor.chain().focus().setImage({ src: result }).run();
      setShowPopup(false);
    };
    reader.readAsDataURL(file);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (file) insertDataUrl(file);
    if (e.target) e.target.value = '';
  };

  const handlePaste = (e: React.ClipboardEvent): void => {
    const items = e.clipboardData?.items;
    if (!items) return;
    for (let i = 0; i < items.length; i++) {
      if (items[i].type.startsWith('image/')) {
        e.preventDefault();
        const file = items[i].getAsFile();
        if (file) insertDataUrl(file);
        return;
      }
    }
  };

  return (
    <div className={styles.tbDropWrap}>
      <Btn title="Insert image" onClick={() => setShowPopup(!showPopup)}>
        {I.image}
      </Btn>
      {showPopup && (
        <div className={styles.imageInsert} ref={popupRef} onPaste={handlePaste}>
          <div className={styles.imageInsertSection}>
            <button className={styles.imageUploadBtn} onClick={() => fileInputRef.current?.click()}>
              Upload from file
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept="image/*"
              style={{ display: 'none' }}
              onChange={handleFileChange}
            />
          </div>
          <div className={styles.imageInsertDivider}>or</div>
          <div className={styles.imageInsertSection}>
            <input
              type="url"
              value={urlValue}
              onChange={e => setUrlValue(e.target.value)}
              placeholder="Paste image URL..."
              onKeyDown={e => {
                if (e.key === 'Enter' && urlValue) {
                  editor.chain().focus().setImage({ src: urlValue }).run();
                  setUrlValue('');
                  setShowPopup(false);
                }
              }}
            />
            <button
              disabled={!urlValue}
              onClick={() => {
                editor.chain().focus().setImage({ src: urlValue }).run();
                setUrlValue('');
                setShowPopup(false);
              }}
            >
              Insert
            </button>
          </div>
          <div className={styles.imageInsertHint}>Tip: You can also paste an image from your clipboard here</div>
        </div>
      )}
    </div>
  );
};

/* ------------------------------------------------------------------ */
/*  Dropdown sub-components                                            */
/* ------------------------------------------------------------------ */

const FontSizeDropdown: React.FC<{ editor: Editor }> = ({ editor }) => {
  const [show, setShow] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) setShow(false);
    };
    if (show) document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [show]);

  const currentSize = (editor.getAttributes('textStyle').fontSize as string) || '';
  let currentNum: number;
  if (currentSize) {
    currentNum = parseInt(currentSize, 10);
  } else if (editor.isActive('heading', { level: 2 })) {
    currentNum = 24;
  } else if (editor.isActive('heading', { level: 3 })) {
    currentNum = 20;
  } else if (editor.isActive('heading', { level: 4 })) {
    currentNum = 18;
  } else {
    currentNum = 16;
  }

  return (
    <div className={styles.tbDropWrap} ref={ref}>
      <button className={styles.sizeBtn} onClick={() => setShow(!show)} onMouseDown={e => e.preventDefault()}>
        {currentNum} <span className={styles.dropArrow}>▾</span>
      </button>
      {show && (
        <div className={styles.fontSizePopup}>
          {FONT_SIZE_OPTIONS.map(size => {
            const isActive = size === currentNum;
            return (
              <button
                key={size}
                className={`${styles.fontSizeOption} ${isActive ? styles.fontSizeActive : ''}`}
                onMouseDown={e => e.preventDefault()}
                onClick={() => {
                  if (size === 16) editor.chain().focus().unsetFontSize().run();
                  else editor.chain().focus().setFontSize(`${size}px`).run();
                  setShow(false);
                }}
              >
                {isActive && <span className={styles.checkmark}>✓</span>}
                <span>{size}</span>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
};

const AlignDropdown: React.FC<{ editor: Editor }> = ({ editor }) => {
  const [show, setShow] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) setShow(false);
    };
    if (show) document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [show]);

  const aligns: { id: string; icon: React.ReactNode; label: string }[] = [
    { id: 'left', icon: I.alignLeft, label: 'Left' },
    { id: 'center', icon: I.alignCenter, label: 'Center' },
    { id: 'right', icon: I.alignRight, label: 'Right' },
    { id: 'justify', icon: I.alignJustify, label: 'Justify' },
  ];
  const current = aligns.find(a => editor.isActive({ textAlign: a.id })) || aligns[0];

  return (
    <div className={styles.tbDropWrap} ref={ref}>
      <Btn title="Text alignment" onClick={() => setShow(!show)}>
        {current.icon}
      </Btn>
      {show && (
        <div className={styles.miniPopup}>
          {aligns.map(a => (
            <Btn key={a.id} active={editor.isActive({ textAlign: a.id })} title={a.label}
              onClick={() => { editor.chain().focus().setTextAlign(a.id).run(); setShow(false); }}>
              {a.icon}
            </Btn>
          ))}
        </div>
      )}
    </div>
  );
};

const ListDropdown: React.FC<{ editor: Editor }> = ({ editor }) => {
  const [show, setShow] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) setShow(false);
    };
    if (show) document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [show]);

  return (
    <div className={styles.tbDropWrap} ref={ref}>
      <Btn title="Lists" active={editor.isActive('bulletList') || editor.isActive('orderedList')} onClick={() => setShow(!show)}>
        {editor.isActive('orderedList') ? I.orderedList : I.bulletList}
      </Btn>
      {show && (
        <div className={styles.miniPopup}>
          <Btn active={editor.isActive('bulletList')} title="Bullet list"
            onClick={() => { editor.chain().focus().toggleBulletList().run(); setShow(false); }}>
            {I.bulletList}
          </Btn>
          <Btn active={editor.isActive('orderedList')} title="Numbered list"
            onClick={() => { editor.chain().focus().toggleOrderedList().run(); setShow(false); }}>
            {I.orderedList}
          </Btn>
        </div>
      )}
    </div>
  );
};

/* ------------------------------------------------------------------ */
/*  TEXT TOOLBAR                                                       */
/* ------------------------------------------------------------------ */

const TextToolbar: React.FC<{ editor: Editor }> = ({ editor }) => {
  const [showFontColor, setShowFontColor] = React.useState(false);
  const [showHighlight, setShowHighlight] = React.useState(false);
  const [showTableGrid, setShowTableGrid] = React.useState(false);
  const [showLinkInput, setShowLinkInput] = React.useState(false);

  const headingLevel = ([2, 3, 4] as const).find(l => editor.isActive('heading', { level: l }));
  const isCodeBlock = editor.isActive('codeBlock');
  const styleValue = isCodeBlock ? 'code' : headingLevel ? `h${headingLevel}` : 'p';
  const currentLineHeight = editor.getAttributes('paragraph').lineHeight as string ||
    editor.getAttributes('heading').lineHeight as string || '';

  return (
    <div className={styles.toolbar}>
      {/* ── Style ── */}
      <div className={styles.tbGroup}>
        <select
          className={styles.tbSelect}
          value={styleValue}
          onChange={e => {
            const v = e.target.value;
            if (v === 'p') editor.chain().focus().clearNodes().setParagraph().run();
            else if (v === 'code') { editor.chain().focus().toggleCodeBlock().run(); return; }
            else editor.chain().focus().toggleHeading({ level: parseInt(v[1], 10) as 2 | 3 | 4 }).run();
            // Clear explicit font-size marks so CSS heading defaults apply
            const { $from } = editor.state.selection;
            const bStart = $from.start($from.depth);
            const bEnd = $from.end($from.depth);
            if (bEnd > bStart) {
              const cursor = $from.pos;
              editor.chain().setTextSelection({ from: bStart, to: bEnd }).unsetFontSize().setTextSelection(cursor).run();
            }
          }}
        >
          <option value="p">Normal</option>
          <option value="h2">Heading 2</option>
          <option value="h3">Heading 3</option>
          <option value="h4">Heading 4</option>
          <option value="code">Monospaced</option>
        </select>
      </div>

      {/* ── Font Size ── */}
      <div className={styles.tbGroup}>
        <FontSizeDropdown editor={editor} />
      </div>

      {/* ── Bold / Italic / Underline / Strikethrough ── */}
      <div className={styles.tbGroup}>
        <Btn active={editor.isActive('bold')} title="Bold (Ctrl+B)" onClick={() => editor.chain().focus().toggleBold().run()}>
          <b>B</b>
        </Btn>
        <Btn active={editor.isActive('italic')} title="Italic (Ctrl+I)" onClick={() => editor.chain().focus().toggleItalic().run()}>
          <i style={{ fontFamily: 'serif' }}>I</i>
        </Btn>
        <Btn active={editor.isActive('underline')} title="Underline (Ctrl+U)" onClick={() => editor.chain().focus().toggleUnderline().run()}>
          <span style={{ textDecoration: 'underline' }}>U</span>
        </Btn>
        <Btn active={editor.isActive('strike')} title="Strikethrough" onClick={() => editor.chain().focus().toggleStrike().run()}>
          <span style={{ textDecoration: 'line-through' }}>S</span>
        </Btn>
      </div>

      {/* ── Font color / Highlight ── */}
      <div className={styles.tbGroup}>
        <div className={styles.tbDropWrap}>
          <Btn title="Font color" onClick={() => { setShowFontColor(!showFontColor); setShowHighlight(false); }}>
            <span className={styles.colorIndicator} style={{ borderBottomColor: (editor.getAttributes('textStyle').color as string) || '#000' }}>A</span>
          </Btn>
          {showFontColor && (
            <ColorPicker colors={COLORS} activeColor={editor.getAttributes('textStyle').color as string}
              onSelect={c => editor.chain().focus().setColor(c).run()} onClear={() => editor.chain().focus().unsetColor().run()}
              onClose={() => setShowFontColor(false)} />
          )}
        </div>
        <div className={styles.tbDropWrap}>
          <Btn title="Highlight" onClick={() => { setShowHighlight(!showHighlight); setShowFontColor(false); }}>
            <span className={styles.colorIndicator} style={{ borderBottomColor: (editor.getAttributes('highlight').color as string) || '#ffff00' }}>{I.highlight}</span>
          </Btn>
          {showHighlight && (
            <ColorPicker colors={COLORS.slice(0, 30)} activeColor={editor.getAttributes('highlight').color as string}
              onSelect={c => editor.chain().focus().toggleHighlight({ color: c }).run()} onClear={() => editor.chain().focus().unsetHighlight().run()}
              onClose={() => setShowHighlight(false)} />
          )}
        </div>
      </div>

      {/* ── Link ── */}
      <div className={styles.tbGroup}>
        <div className={styles.tbDropWrap}>
          <Btn active={editor.isActive('link')} title="Insert link (Ctrl+K)" onClick={() => setShowLinkInput(!showLinkInput)}>
            {I.link}
          </Btn>
          {showLinkInput && (
            <LinkInput
              initialUrl={editor.getAttributes('link').href as string}
              onSubmit={url => editor.chain().focus().extendMarkRange('link').setLink({ href: url, target: '_blank' }).run()}
              onRemove={() => editor.chain().focus().extendMarkRange('link').unsetLink().run()}
              onClose={() => setShowLinkInput(false)}
            />
          )}
        </div>
      </div>

      {/* ── Align / Lists ── */}
      <div className={styles.tbGroup}>
        <AlignDropdown editor={editor} />
        <ListDropdown editor={editor} />
      </div>

      {/* ── Superscript / Subscript ── */}
      <div className={styles.tbGroup}>
        <Btn active={editor.isActive('superscript')} title="Superscript" onClick={() => editor.chain().focus().toggleSuperscript().run()}>
          <span className={styles.scriptBtn}>x<sup>2</sup></span>
        </Btn>
        <Btn active={editor.isActive('subscript')} title="Subscript" onClick={() => editor.chain().focus().toggleSubscript().run()}>
          <span className={styles.scriptBtn}>x<sub>2</sub></span>
        </Btn>
      </div>

      {/* ── Indent ── */}
      <div className={styles.tbGroup}>
        <Btn title="Decrease indent" onClick={() => editor.chain().focus().decreaseIndent().run()}>
          {I.outdent}
        </Btn>
        <Btn title="Increase indent" onClick={() => editor.chain().focus().increaseIndent().run()}>
          {I.indent}
        </Btn>
      </div>

      {/* ── Line spacing ── */}
      <div className={styles.tbGroup}>
        <select className={styles.tbSelect} value={currentLineHeight}
          onChange={e => {
            if (!e.target.value) editor.chain().focus().unsetLineHeight().run();
            else editor.chain().focus().setLineHeight(e.target.value).run();
          }}>
          <option value="">Spacing</option>
          {LINE_HEIGHTS.map(lh => <option key={lh.value} value={lh.value}>{lh.label}</option>)}
        </select>
      </div>

      {/* ── Image / Table ── */}
      <div className={styles.tbGroup}>
        <ImageInsert editor={editor} />
        <div className={styles.tbDropWrap}>
          <Btn title="Insert table" onClick={() => setShowTableGrid(!showTableGrid)}>
            {I.table}
          </Btn>
          {showTableGrid && (
            <TableGridSelector onSelect={(r, c) => editor.chain().focus().insertTable({ rows: r, cols: c, withHeaderRow: true }).run()}
              onClose={() => setShowTableGrid(false)} />
          )}
        </div>
      </div>

      {/* ── Clear formatting ── */}
      <div className={styles.tbGroup}>
        <Btn title="Clear formatting" onClick={() => editor.chain().focus().clearNodes().unsetAllMarks().run()}>
          {I.clearFormat}
        </Btn>
      </div>
    </div>
  );
};

/* ------------------------------------------------------------------ */
/*  TABLE TOOLBAR                                                      */
/* ------------------------------------------------------------------ */

// Helper: get positions of all selected cells
// Reads from lastSelectedCellPositions tracked by CellTracker plugin
// — these are captured INSIDE the editor before focus/blur resets selection
function getSelectedCellPositions(editor: Editor): number[] {
  const cellNodeTypes = ['tableCell', 'tableHeader'];

  // Use plugin-tracked positions — most reliable, survives focus loss
  if (lastSelectedCellPositions.length > 0) {
    return [...lastSelectedCellPositions];
  }

  // Fallback: single cell at cursor
  const positions: number[] = [];
  const { $from } = editor.state.selection;
  for (let d = $from.depth; d >= 0; d--) {
    if (cellNodeTypes.includes($from.node(d).type.name)) {
      positions.push($from.before(d));
      break;
    }
  }
  return positions;
}

const TableToolbar: React.FC<{ editor: Editor }> = ({ editor }) => {
  const [showBorderColor, setShowBorderColor] = React.useState(false);
  const [showCellBg, setShowCellBg] = React.useState(false);
  // Store cell positions at mousedown time — before focus changes
  const cellPositionsRef = React.useRef<number[]>([]);

  const applyAttr = (attr: string, value: unknown): void => {
    const positions = cellPositionsRef.current.length > 0
      ? cellPositionsRef.current
      : getSelectedCellPositions(editor);

    if (positions.length === 0) return;

    const { state, view } = editor;
    const tr = state.tr;
    positions.forEach(pos => {
      const node = state.doc.nodeAt(pos);
      if (node) tr.setNodeMarkup(pos, undefined, { ...node.attrs, [attr]: value });
    });
    view.dispatch(tr);
  };

  return (
    <div
      className={styles.tableToolbar}
      onMouseDown={(e) => {
        // Capture selected cell positions BEFORE any focus/blur happens
        cellPositionsRef.current = getSelectedCellPositions(editor);
        // Prevent blur for buttons — but allow select dropdowns to work
        const tag = (e.target as HTMLElement).tagName.toLowerCase();
        if (tag !== 'select' && tag !== 'option') {
          e.preventDefault();
        }
      }}
    >
      {/* ── Add row/col ── */}
      <div className={styles.tbGroup}>
        <Btn title="Add row above" onClick={() => editor.chain().focus().addRowBefore().run()}>{I.addRowAbove}</Btn>
        <Btn title="Add row below" onClick={() => editor.chain().focus().addRowAfter().run()}>{I.addRowBelow}</Btn>
        <Btn title="Add column left" onClick={() => editor.chain().focus().addColumnBefore().run()}>{I.addColLeft}</Btn>
        <Btn title="Add column right" onClick={() => editor.chain().focus().addColumnAfter().run()}>{I.addColRight}</Btn>
      </div>

      {/* ── Delete ── */}
      <div className={styles.tbGroup}>
        <Btn title="Delete row" onClick={() => editor.chain().focus().deleteRow().run()}>{I.deleteRow}</Btn>
        <Btn title="Delete column" onClick={() => editor.chain().focus().deleteColumn().run()}>{I.deleteCol}</Btn>
        <Btn title="Delete table" onClick={() => editor.chain().focus().deleteTable().run()} className={styles.tbDanger}>{I.deleteTable}</Btn>
      </div>

      {/* ── Merge / Split ── */}
      <div className={styles.tbGroup}>
        <Btn title="Merge cells" disabled={!editor.can().mergeCells()} onClick={() => editor.chain().focus().mergeCells().run()}>{I.mergeCells}</Btn>
        <Btn title="Split cell" disabled={!editor.can().splitCell()} onClick={() => editor.chain().focus().splitCell().run()}>{I.splitCell}</Btn>
      </div>

      {/* ── Border style ── */}
      <div className={styles.tbGroup}>
        <select className={styles.tbSelect} defaultValue="solid"
          onChange={e => applyAttr('borderStyle', e.target.value)}>
          {BORDER_STYLES.map(b => <option key={b.value} value={b.value}>{b.label}</option>)}
        </select>
      </div>

      {/* ── Border color / Cell background ── */}
      <div className={styles.tbGroup}>
        <div className={styles.tbDropWrap}>
          <Btn title="Border color" onClick={() => { setShowBorderColor(!showBorderColor); setShowCellBg(false); }}>
            <span className={styles.colorIndicator} style={{ borderBottomColor: '#666' }}>B</span>
          </Btn>
          {showBorderColor && (
            <ColorPicker colors={COLORS.slice(0, 20)}
              onSelect={c => { applyAttr('borderColor', c); setShowBorderColor(false); }}
              onClose={() => setShowBorderColor(false)} />
          )}
        </div>
        <div className={styles.tbDropWrap}>
          <Btn title="Cell background" onClick={() => { setShowCellBg(!showCellBg); setShowBorderColor(false); }}>
            <span className={styles.highlightIndicator}>bg</span>
          </Btn>
          {showCellBg && (
            <ColorPicker colors={COLORS}
              onSelect={c => { applyAttr('backgroundColor', c); setShowCellBg(false); }}
              onClear={() => { applyAttr('backgroundColor', null); setShowCellBg(false); }}
              onClose={() => setShowCellBg(false)} />
          )}
        </div>
      </div>

      {/* ── Cell padding ── */}
      <div className={styles.tbGroup}>
        <select className={styles.tbSelect} defaultValue="8px 12px"
          onChange={e => applyAttr('cellPadding', e.target.value)}>
          {CELL_PADDINGS.map(p => <option key={p.value} value={p.value}>{p.label}</option>)}
        </select>
      </div>

      {/* ── Paragraph before/after ── */}
      <div className={styles.tbGroup}>
        <Btn title="Add line before table" onClick={() => addLineBefore(editor)}>{I.paragraphBefore}</Btn>
        <Btn title="Add line after table" onClick={() => addLineAfter(editor)}>{I.paragraphAfter}</Btn>
      </div>
    </div>
  );
};

/* ------------------------------------------------------------------ */
/*  Main Editor Component                                              */
/* ------------------------------------------------------------------ */

export interface ITipTapEditorProps {
  content: string;
  onChange: (html: string) => void;
  placeholder?: string;
}

export const TipTapEditor: React.FC<ITipTapEditorProps> = ({ content, onChange, placeholder }) => {
  const lastContentRef = React.useRef(content);
  const [, forceUpdate] = React.useState(0);
  const wrapperRef = React.useRef<HTMLDivElement>(null);

  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        heading: { levels: [2, 3, 4] },
      }),
      TextStyle,
      FontSize,
      Color,
      Highlight.configure({ multicolor: true }),
      Underline,
      Superscript,
      Subscript,
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Link.configure({ openOnClick: false, HTMLAttributes: { target: '_blank', rel: 'noopener noreferrer' } }),
      ResizableImage,
      Table.configure({ resizable: true, cellMinWidth: 50 }),
      TableRow,
      CustomTableCell,
      CustomTableHeader,
      BulkCellAttribute,
      CellTracker,
      Placeholder.configure({ placeholder: placeholder || 'Start writing your content here…' }),
      LineHeight,
      Indent,
    ],
    content,
    onUpdate: ({ editor: ed }) => {
      const html = ed.getHTML();
      lastContentRef.current = html;
      onChange(html);
    },
    onSelectionUpdate: () => {
      // Force re-render to update toolbar state
      forceUpdate(n => n + 1);
    },
    onTransaction: () => {
      forceUpdate(n => n + 1);
    },
  });

  // Sync external content changes (e.g. from property pane)
  React.useEffect(() => {
    if (editor && content !== lastContentRef.current) {
      editor.commands.setContent(content, { emitUpdate: false });
      lastContentRef.current = content;
    }
  }, [content, editor]);

  if (!editor) return null;

  const isInTable = editor.isActive('table');

  // Calculate floating table toolbar position
  let tableToolbarStyle: React.CSSProperties | undefined;
  if (isInTable && wrapperRef.current) {
    const pos = findTablePos(editor);
    if (pos !== null) {
      const dom = editor.view.nodeDOM(pos);
      if (dom && dom instanceof HTMLElement) {
        let offsetTop = 0;
        let el: HTMLElement | null = dom;
        while (el && el !== wrapperRef.current) {
          offsetTop += el.offsetTop;
          el = el.offsetParent as HTMLElement | null;
        }
        tableToolbarStyle = {
          position: 'absolute' as const,
          top: offsetTop,
          left: 0,
          right: 0,
          transform: 'translateY(-100%)',
          zIndex: 11,
        };
      }
    }
  }

  return (
    <div className={styles.editorWrapper} ref={wrapperRef}>
      <TextToolbar editor={editor} />
      <EditorContent editor={editor} className={styles.editorContent} />
      {isInTable && tableToolbarStyle && (
        <div style={tableToolbarStyle}>
          <TableToolbar editor={editor} />
        </div>
      )}
    </div>
  );
};
