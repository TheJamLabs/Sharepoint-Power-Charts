import * as React from 'react';
import { DataColumn, DataRow, TableBlockConfig } from '../../data/types';
import { aggregate, groupBy } from '../../data/query';
import { formatAggregated, formatCell } from '../../utils/format';
import styles from '../SharepointPowerCharts.module.scss';

export interface TableBlockProps {
  config: TableBlockConfig;
  /** Source column metadata (for friendly headers / fallback columns). */
  columns: DataColumn[];
  /** Rows already filtered (global + block) for this source. */
  rows: DataRow[];
}

/**
 * Renders a data table. With `groupBy` + `subtotalColumn` set, rows are grouped
 * with a subtotal line per group; otherwise rows are shown flat. `limitRows`
 * caps the number of group headers (or flat rows).
 */
export default class TableBlock extends React.Component<TableBlockProps, {}> {
  private resolvedColumns(): string[] {
    const { config, columns } = this.props;
    if (config.columns && config.columns.length > 0) {
      return config.columns;
    }
    return columns.slice(0, 4).map((c) => c.internalName);
  }

  private columnLabel(internalName: string): string {
    const { columns } = this.props;
    const meta = columns.filter((c) => c.internalName === internalName)[0];
    return meta ? meta.displayName : internalName;
  }

  private isNumericColumn(internalName: string): boolean {
    const { columns } = this.props;
    const meta = columns.filter((c) => c.internalName === internalName)[0];
    return !!meta && meta.kind === 'number';
  }

  public render(): React.ReactElement<{}> {
    const { config, rows } = this.props;
    const cols = this.resolvedColumns();
    if (cols.length === 0 || rows.length === 0) {
      return <div className={styles.chartEmpty}>No rows to display.</div>;
    }

    const hasGrouping = !!config.groupBy && cols.indexOf(config.groupBy as string) >= 0;
    const headers = hasGrouping ? cols.filter((c) => c !== config.groupBy) : cols;

    return (
      <div style={{ overflowX: 'auto' }}>
        <table className={styles.table}>
          <thead>
            <tr>
              {headers.map((h, i) => (
                <th key={h + i} className={this.isNumericColumn(h) ? styles.thNum : undefined}>
                  {this.columnLabel(h)}
                </th>
              ))}
            </tr>
          </thead>
          {hasGrouping ? this.renderGrouped(rows, cols, config) : this.renderFlat(rows, cols)}
        </table>
      </div>
    );
  }

  private renderFlat(rows: DataRow[], cols: string[]): React.ReactElement<{}> {
    const limit = this.props.config.limitRows || 1000;
    const shown = rows.slice(0, limit);
    return (
      <tbody>
        {shown.map((row, ri) => (
          <tr key={ri}>
            {cols.map((c) => (
              <td key={c} className={this.isNumericColumn(c) ? styles.tableCellNumber : undefined}>
                {formatCell(row[c])}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    );
  }

  private renderGrouped(
    rows: DataRow[],
    cols: string[],
    config: TableBlockConfig
  ): React.ReactElement<{}> {
    const groupCol = config.groupBy as string;
    const dataCols = cols.filter((c) => c !== groupCol);
    const limit = config.limitRows || 1000;
    const groups = groupBy(rows, groupCol).slice(0, limit);
    const subtotalCol = config.subtotalColumn;
    const subtotalAgg = config.subtotalAgg || 'sum';

    return (
      <tbody>
        {groups.map((g, gi) => {
          const subtotal =
            subtotalCol && dataCols.indexOf(subtotalCol) >= 0
              ? aggregate(g.rows, subtotalCol, subtotalAgg)
              : undefined;
          return (
            <React.Fragment key={gi}>
              <tr className={styles.groupRow}>
                <td colSpan={cols.length}>{g.label}</td>
              </tr>
              {g.rows.map((row, ri) => (
                <tr key={ri}>
                  {dataCols.map((c) => (
                    <td key={c} className={this.isNumericColumn(c) ? styles.tableCellNumber : undefined}>
                      {formatCell(row[c])}
                    </td>
                  ))}
                </tr>
              ))}
              {subtotal !== undefined && (
                <tr className={styles.subtotalRow}>
                  <td colSpan={dataCols.length - 1}>Subtotal</td>
                  <td className={styles.tableCellNumber}>{formatAggregated(subtotal, subtotalAgg)}</td>
                </tr>
              )}
            </React.Fragment>
          );
        })}
      </tbody>
    );
  }
}