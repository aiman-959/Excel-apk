import {
  Bold, Italic, Underline, Strikethrough, AlignLeft, AlignCenter, AlignRight, WrapText, Merge, Split, Square, Grid3x3, PanelTop, PanelBottom, SquareDashed,
  Hash, Highlighter, Eraser, Search, ArrowDownAZ, ArrowUpZA, Filter, FilterX, Snowflake, BarChart3, LineChart, PieChart, Table2, Sigma, FileSpreadsheet,
  Code2, Workflow, Gauge, Rows3, Columns3, Trash2, Undo2, Redo2, ListX, Info,
} from 'lucide-react';
import { Tool } from './Chrome';
import { data, edit, fmt, sheets } from './actions';
import { guard, ui, type ToolTab } from './store';
import { useSelection } from './hooks';

const COLORS = ['#14231a', '#4a5a50', '#b42318', '#e8772e', '#fbbf24', '#15803d', '#1e88c7', '#14532d', '#7c3aed', '#ffffff', '#fde2e4', '#ffe8cc', '#fef3c7', '#dcfce7', '#d0ebff', '#ede9fe'];

const run = (fn: () => unknown) => () => guard(fn);

export function ToolPanel({ tab }: { tab: ToolTab }) {
  const sel = useSelection();
  switch (tab) {
    case 'edit':
      return (
        <>
          <div className="group">
            <h4>Edit</h4>
            <div className="btn-grid">
              <Tool icon={<Undo2 />} label="Undo" onClick={edit.undo} />
              <Tool icon={<Redo2 />} label="Redo" onClick={edit.redo} />
              <Tool icon={<Search />} label="Find & replace" testId="tool-find" onClick={() => ui.open('find')} />
              <Tool icon={<Eraser />} label="Clear" onClick={run(edit.clear)} />
            </div>
          </div>
          <div className="group">
            <h4>Rows & columns</h4>
            <div className="btn-grid">
              <Tool icon={<Rows3 />} label="Row above" onClick={run(edit.insertRowAbove)} />
              <Tool icon={<Rows3 />} label="Row below" onClick={run(edit.insertRowBelow)} />
              <Tool icon={<Columns3 />} label="Column left" onClick={run(edit.insertColLeft)} />
              <Tool icon={<Columns3 />} label="Column right" onClick={run(edit.insertColRight)} />
              <Tool icon={<Trash2 />} label="Delete rows" onClick={run(edit.deleteRows)} />
              <Tool icon={<Trash2 />} label="Delete cols" onClick={run(edit.deleteCols)} />
            </div>
          </div>
        </>
      );
    case 'format':
      return (
        <>
          <div className="group">
            <h4>Text</h4>
            <div className="btn-grid">
              <Tool icon={<Bold />} label="Bold" on={sel.bold} testId="fmt-bold" onClick={run(fmt.bold)} />
              <Tool icon={<Italic />} label="Italic" on={sel.italic} onClick={run(fmt.italic)} />
              <Tool icon={<Underline />} label="Underline" onClick={run(fmt.underline)} />
              <Tool icon={<Strikethrough />} label="Strike" onClick={run(fmt.strike)} />
              <Tool icon={<AlignLeft />} label="Left" onClick={run(() => fmt.align('left'))} />
              <Tool icon={<AlignCenter />} label="Center" onClick={run(() => fmt.align('center'))} />
              <Tool icon={<AlignRight />} label="Right" onClick={run(() => fmt.align('right'))} />
              <Tool icon={<WrapText />} label="Wrap" onClick={run(fmt.wrap)} />
            </div>
          </div>
          <div className="group">
            <h4>Fill colour</h4>
            <div className="swatches">
              <button className="swatch none" aria-label="No fill" onClick={run(() => fmt.fill(null))} />
              {COLORS.slice(1).map((c) => (
                <button key={c} className="swatch" style={{ background: c }} aria-label={'Fill ' + c} data-testid={'fill-' + c} onClick={run(() => fmt.fill(c))} />
              ))}
            </div>
          </div>
          <div className="group">
            <h4>Text colour</h4>
            <div className="swatches">
              {COLORS.slice(0, 9).map((c) => (
                <button key={c} className="swatch" style={{ background: c }} aria-label={'Text ' + c} onClick={run(() => fmt.color(c))} />
              ))}
            </div>
          </div>
          <div className="group">
            <h4>Cells</h4>
            <div className="btn-grid">
              <Tool icon={<Merge />} label="Merge" testId="fmt-merge" onClick={run(fmt.merge)} />
              <Tool icon={<Split />} label="Unmerge" onClick={run(fmt.unmerge)} />
              <Tool icon={<Grid3x3 />} label="All borders" testId="fmt-border-all" onClick={run(() => fmt.border('all'))} />
              <Tool icon={<Square />} label="Outline" onClick={run(() => fmt.border('outside', true))} />
              <Tool icon={<PanelTop />} label="Top" onClick={run(() => fmt.border('top'))} />
              <Tool icon={<PanelBottom />} label="Bottom" onClick={run(() => fmt.border('bottom'))} />
              <Tool icon={<SquareDashed />} label="No border" onClick={run(() => fmt.border('none'))} />
              <Tool icon={<Eraser />} label="Clear format" onClick={run(fmt.clearFormat)} />
            </div>
          </div>
          <div className="group">
            <h4>Numbers & rules</h4>
            <div className="btn-grid">
              <Tool icon={<Hash />} label="Number format" testId="tool-numfmt" onClick={() => ui.open('numfmt')} />
              <Tool icon={<Highlighter />} label="Conditional" testId="tool-cf" onClick={() => ui.open('cf')} />
            </div>
          </div>
        </>
      );
    case 'insert':
      return (
        <div className="group">
          <h4>Insert</h4>
          <div className="btn-grid">
            <Tool icon={<BarChart3 />} label="Bar chart" testId="tool-chart-bar" onClick={() => ui.open('chart', 'bar')} />
            <Tool icon={<LineChart />} label="Line chart" onClick={() => ui.open('chart', 'line')} />
            <Tool icon={<PieChart />} label="Pie chart" onClick={() => ui.open('chart', 'pie')} />
            <Tool icon={<Table2 />} label="Pivot table" testId="tool-pivot" onClick={() => ui.open('pivot')} />
            <Tool icon={<Sigma />} label="Function" onClick={() => ui.open('functions')} />
            <Tool icon={<FileSpreadsheet />} label="New sheet" onClick={run(sheets.add)} />
          </div>
        </div>
      );
    case 'data':
      return (
        <>
          <div className="group">
            <h4>Sort & filter</h4>
            <div className="btn-grid">
              <Tool icon={<ArrowDownAZ />} label="Sort A→Z" testId="tool-sort-asc" onClick={run(() => data.sort(true))} />
              <Tool icon={<ArrowUpZA />} label="Sort Z→A" onClick={run(() => data.sort(false))} />
              <Tool icon={<Filter />} label="Filter on/off" testId="tool-filter" onClick={run(() => ui.toast(data.toggleFilter() ? 'Filter on — tap a header arrow' : 'Filter removed'))} />
              <Tool icon={<FilterX />} label="Clear filters" onClick={run(data.clearFilters)} />
            </div>
          </div>
          <div className="group">
            <h4>Freeze panes</h4>
            <div className="btn-grid">
              <Tool icon={<Snowflake />} label="Top row" testId="tool-freeze-row" onClick={run(data.freezeTopRow)} />
              <Tool icon={<Snowflake />} label="First column" onClick={run(data.freezeFirstCol)} />
              <Tool icon={<Snowflake />} label="At selection" onClick={run(data.freezeAtSelection)} />
              <Tool icon={<ListX />} label="Unfreeze" onClick={run(data.unfreeze)} />
            </div>
          </div>
          <div className="group">
            <h4>Analyse</h4>
            <div className="btn-grid">
              <Tool icon={<Table2 />} label="Pivot table" onClick={() => ui.open('pivot')} />
              <Tool icon={<Workflow />} label="Transform" testId="tool-transform" onClick={() => ui.open('transform')} />
            </div>
          </div>
        </>
      );
    case 'tools':
      return (
        <div className="group">
          <h4>Tools</h4>
          <div className="btn-grid">
            <Tool icon={<Code2 />} label="Macros" testId="tool-macros" onClick={() => ui.open('macros')} />
            <Tool icon={<Workflow />} label="Transform" onClick={() => ui.open('transform')} />
            <Tool icon={<BarChart3 />} label="Charts" onClick={() => ui.open('chart')} />
            <Tool icon={<Gauge />} label="Device & perf" onClick={() => ui.open('about')} />
            <Tool icon={<Info />} label="About" onClick={() => ui.open('about')} />
          </div>
        </div>
      );
  }
}
