import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, Link } from '@fluentui/react';
import { HELP_MARKDOWN } from './helpContent.g';

/**
 * In-app "?" help: renders docs/configuration.md (embedded as a generated
 * string constant) inside the property pane, so the configuration reference is
 * available without opening the repository. A deliberately small markdown
 * subset is supported — exactly the constructs the reference file uses:
 * ATX headings, pipe tables, fenced code blocks, bullet/numbered lists,
 * blockquotes, horizontal rules, and inline code / bold / links.
 */

// ---------------------------------------------------------------------------
// Markdown subset parser
// ---------------------------------------------------------------------------

type InlineNode = { kind: 'text'; text: string } | { kind: 'code'; text: string } | { kind: 'bold'; text: string };

interface HeadingBlock {
  kind: 'heading';
  level: 1 | 2 | 3;
  id: string;
  text: string;
}
interface ParagraphBlock {
  kind: 'p';
  text: string;
}
interface ListBlock {
  kind: 'ul' | 'ol';
  items: string[];
}
interface TableBlock {
  kind: 'table';
  header: string[];
  rows: string[][];
}
interface CodeBlock {
  kind: 'code';
  text: string;
}
interface QuoteBlock {
  kind: 'quote';
  text: string;
}
interface HrBlock {
  kind: 'hr';
}

type MarkdownBlock = HeadingBlock | ParagraphBlock | ListBlock | TableBlock | CodeBlock | QuoteBlock | HrBlock;

function parseInline(text: string): InlineNode[] {
  const out: InlineNode[] = [];
  let plain = '';
  let i = 0;
  const flush = (): void => {
    if (plain !== '') {
      out.push({ kind: 'text', text: plain });
      plain = '';
    }
  };
  while (i < text.length) {
    if (text.startsWith('**', i)) {
      const end = text.indexOf('**', i + 2);
      if (end === -1) {
        plain += text.slice(i);
        break;
      }
      flush();
      out.push({ kind: 'bold', text: text.slice(i + 2, end) });
      i = end + 2;
      continue;
    }
    if (text[i] === '`') {
      const end = text.indexOf('`', i + 1);
      if (end === -1) {
        plain += text.slice(i);
        break;
      }
      flush();
      out.push({ kind: 'code', text: text.slice(i + 1, end) });
      i = end + 1;
      continue;
    }
    if (text[i] === '[') {
      // Markdown link: render the label only (no navigation inside the dialog).
      const close = text.indexOf('](', i);
      const end = close === -1 ? -1 : text.indexOf(')', close + 2);
      if (end !== -1) {
        flush();
        out.push({ kind: 'text', text: text.slice(i + 1, close) });
        i = end + 1;
        continue;
      }
    }
    plain += text[i];
    i++;
  }
  flush();
  return out;
}

function slugify(text: string): string {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function parseMarkdown(markdown: string): MarkdownBlock[] {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n');
  const blocks: MarkdownBlock[] = [];
  let i = 0;

  const isTableRow = (l: string): boolean => l.trim().startsWith('|') && l.split('|').length >= 3;
  const isSeparatorRow = (l: string): boolean => {
    const t = l
      .trim()
      .replace(/^\|/, '')
      .replace(/\|$/, '');
    return t !== '' && /^[\s:|-]+$/.test(t) && t.indexOf('-') !== -1;
  };
  const isListItem = (l: string): string | null => {
    const m = l.match(/^\s*([-*]|\d+[.)])\s+(.*)$/);
    return m ? m[2] : null;
  };

  while (i < lines.length) {
    const trimmed = lines[i].trim();
    if (trimmed === '') {
      i++;
      continue;
    }
    if (trimmed.startsWith('```')) {
      const buf: string[] = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith('```')) {
        buf.push(lines[i]);
        i++;
      }
      i++; // closing fence
      blocks.push({ kind: 'code', text: buf.join('\n') });
      continue;
    }
    const heading = trimmed.match(/^(#{1,3})\s+(.*)$/);
    if (heading) {
      blocks.push({
        kind: 'heading',
        level: heading[1].length as 1 | 2 | 3,
        id: slugify(heading[2]),
        text: heading[2]
      });
      i++;
      continue;
    }
    if (/^-{3,}\s*$/.test(trimmed)) {
      blocks.push({ kind: 'hr' });
      i++;
      continue;
    }
    if (isTableRow(trimmed)) {
      const raw: string[] = [];
      while (i < lines.length && isTableRow(lines[i])) {
        raw.push(lines[i]);
        i++;
      }
      const cells = (l: string): string[] =>
        l
          .trim()
          .replace(/^\|/, '')
          .replace(/\|$/, '')
          .split('|')
          .map((c) => c.trim());
      blocks.push({
        kind: 'table',
        header: cells(raw[0]),
        rows: raw.slice(1).filter((r) => !isSeparatorRow(r)).map(cells)
      });
      continue;
    }
    if (trimmed.startsWith('>')) {
      const buf: string[] = [];
      while (i < lines.length && lines[i].trim().startsWith('>')) {
        buf.push(lines[i].trim().replace(/^>\s?/, ''));
        i++;
      }
      blocks.push({ kind: 'quote', text: buf.join(' ') });
      continue;
    }
    const item = isListItem(trimmed);
    if (item !== null) {
      const ordered = /^\s*\d+[.)]\s/.test(lines[i]);
      const items: string[] = [item];
      i++;
      while (i < lines.length) {
        const next = isListItem(lines[i]);
        const nextOrdered = /^\s*\d+[.)]\s/.test(lines[i]);
        if (next === null || nextOrdered !== ordered) {
          break;
        }
        items.push(next);
        i++;
      }
      blocks.push({ kind: ordered ? 'ol' : 'ul', items });
      continue;
    }
    // Plain paragraph: consume until a blank line or a structural construct.
    const buf: string[] = [trimmed];
    i++;
    while (
      i < lines.length &&
      lines[i].trim() !== '' &&
      !lines[i].trim().startsWith('```') &&
      !/^(#{1,3})\s/.test(lines[i]) &&
      !isTableRow(lines[i]) &&
      !lines[i].trim().startsWith('>') &&
      !/^\s*([-*]|\d+[.)])\s/.test(lines[i]) &&
      !/^-{3,}\s*$/.test(lines[i].trim())
    ) {
      buf.push(lines[i].trim());
      i++;
    }
    blocks.push({ kind: 'p', text: buf.join(' ') });
  }
  return blocks;
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

const CODE_INLINE: React.CSSProperties = {
  background: '#f3f2f1',
  padding: '1px 4px',
  borderRadius: 3,
  fontFamily: 'Consolas, "Courier New", monospace',
  fontSize: 12
};

const CODE_BLOCK: React.CSSProperties = {
  background: '#f3f2f1',
  padding: '8px 10px',
  borderRadius: 4,
  overflowX: 'auto',
  fontFamily: 'Consolas, "Courier New", monospace',
  fontSize: 12,
  whiteSpace: 'pre-wrap',
  margin: '8px 0'
};

const TH_TD: React.CSSProperties = {
  border: '1px solid #edebe9',
  padding: '6px 8px',
  textAlign: 'left',
  verticalAlign: 'top',
  fontSize: 12.5
};

const BODY_STYLE: React.CSSProperties = { fontSize: 13, lineHeight: 1.55 };

function renderInline(nodes: InlineNode[], keyPrefix: string): React.ReactNode {
  return nodes.map((n, idx) => {
    const key = `${keyPrefix}-${idx}`;
    if (n.kind === 'code') {
      return (
        <code key={key} style={CODE_INLINE}>
          {n.text}
        </code>
      );
    }
    if (n.kind === 'bold') {
      return <strong key={key}>{n.text}</strong>;
    }
    return <span key={key}>{n.text}</span>;
  });
}

/** The heading-2 sections used by the sticky table of contents. */
const SECTIONS: HeadingBlock[] = parseMarkdown(HELP_MARKDOWN).filter(
  (b): b is HeadingBlock => b.kind === 'heading' && b.level === 2
);

/** The rendered body (the document title and TOC scaffolding excluded). */
const BODY: MarkdownBlock[] = parseMarkdown(HELP_MARKDOWN).filter(
  (b): b is MarkdownBlock => !(b.kind === 'heading' && b.level === 1)
);

/** Sticky table of contents: jumps to each `##` section of the reference. */
const TOC: React.ReactNode = (
  <div style={{ position: 'sticky', top: 0, background: '#ffffff', zIndex: 1, borderBottom: '1px solid #edebe9', padding: '6px 0 8px', marginBottom: 4 }}>
    {SECTIONS.map((s) => (
      <Link
        key={s.id}
        styles={{ root: { fontSize: 12, marginRight: 12 } }}
        onClick={() => {
          const el = document.getElementById(s.id);
          if (el) {
            el.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }
        }}
      >
        {s.text}
      </Link>
    ))}
  </div>
);

function renderBlocks(blocks: MarkdownBlock[]): React.ReactNode {
  return blocks.map((b, idx) => {
    const key = `b${idx}`;
    switch (b.kind) {
      case 'heading':
        if (b.level === 1) {
          return null;
        }
        return (
          <h2
            key={key}
            id={b.id}
            style={{ fontSize: b.level === 2 ? 15 : 14, fontWeight: 600, margin: '16px 0 6px' }}
          >
            {b.text}
          </h2>
        );
      case 'p':
        return (
          <p key={key} style={{ margin: '6px 0' }}>
            {renderInline(parseInline(b.text), key)}
          </p>
        );
      case 'ul':
      case 'ol':
        return (
          <div key={key}>
            {b.items.map((item, i2) => (
              <div key={i2} style={{ margin: '3px 0', paddingLeft: 18, display: 'flex' }}>
                <span style={{ width: 18, flexShrink: 0, color: '#605e5c' }}>{b.kind === 'ol' ? `${i2 + 1}.` : '•'}</span>
                <span>{renderInline(parseInline(item), `${key}-${i2}`)}</span>
              </div>
            ))}
          </div>
        );
      case 'table':
        return (
          <div key={key} style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', margin: '8px 0' }}>
              <thead>
                <tr>
                  {b.header.map((cell, c) => (
                    <th key={c} style={{ ...TH_TD, background: '#faf9f8', fontWeight: 600 }}>
                      {renderInline(parseInline(cell), `${key}-h${c}`)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {b.rows.map((row, r) => (
                  <tr key={r}>
                    {row.map((cell, c) => (
                      <td key={c} style={TH_TD}>
                        {renderInline(parseInline(cell), `${key}-${r}-${c}`)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
      case 'code':
        return (
          <pre key={key} style={CODE_BLOCK}>
            <code>{b.text}</code>
          </pre>
        );
      case 'quote':
        return (
          <blockquote
            key={key}
            style={{ borderLeft: '3px solid #0078d4', background: 'rgba(0, 120, 212, 0.06)', margin: '8px 0', padding: '6px 10px' }}
          >
            {renderInline(parseInline(b.text), key)}
          </blockquote>
        );
      case 'hr':
        return <div key={key} style={{ borderTop: '1px solid #edebe9', margin: '14px 0' }} />;
      default:
        return null;
    }
  });
}

export interface HelpDialogProps {
  open: boolean;
  onDismiss: () => void;
}

/** "?" help pane: the configuration reference rendered inside the property pane. */
export default function HelpDialog(props: HelpDialogProps): React.ReactElement {
  return (
    <Dialog
      hidden={!props.open}
      onDismiss={props.onDismiss}
      minWidth={680}
      maxWidth={900}
      dialogContentProps={{
        title: 'Configuration help',
        subText: 'The same reference as docs/configuration.md in the repository.'
      }}
    >
      <div style={{ maxHeight: '60vh', overflowY: 'auto', paddingRight: 8 }}>
        <div style={BODY_STYLE}>
          {SECTIONS.length > 0 ? TOC : null}
          {renderBlocks(BODY)}
        </div>
      </div>
      <DialogFooter>
        <DefaultButton text="Close" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
}