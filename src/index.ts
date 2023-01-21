import { Workbook, Worksheet } from 'exceljs';
import { flatten, head, sortBy, startCase, unionBy } from 'lodash';
import { tmpdir } from 'os';
import { HeaderRow, ValueType } from './types';

// TODO: add support for array of objects

export class JsonToExcel {
  export(opts: {
    data: any[];
    excludeFields?: string[];
    workbook?: Workbook;
    worksheet?: Worksheet;
  }): Workbook {
    if (!Array.isArray(opts.data)) {
      opts.data = [opts.data];
    }

    const workbook = opts.workbook ?? new Workbook();
    if (opts.workbook == null) {
      workbook.created = new Date();
      workbook.modified = new Date();
      workbook.lastPrinted = new Date();
    }

    // If no data is provided or is empty, exit immediately
    if (opts.data.length == 0) {
      return workbook;
    }

    const headerRows = this.generateHeaderRow({
      root: '',
      data: opts.data,
      excludeFields: opts.excludeFields ?? [],
    });

    // console.log(headerRows);

    const transformedHeaders = this._transformHeadersForExport(headerRows)
      .filter((i) => i.hidden == false)
      .map((i) => {
        return {
          header: i.title as string,
          key: i.path,
          width: 20,
        };
      });

    // Add a worksheet
    const worksheet = opts.worksheet ?? workbook.addWorksheet('Sheet 1');

    // Add column headers and define column keys and widths
    // Note: these column structures are a workbook-building convenience only,
    // apart from the column width, they will not be fully persisted.
    worksheet.columns = transformedHeaders;

    opts.data.forEach((i) => {
      const row = this.createRow({ item: i, headers: headerRows, root: '' });

      worksheet.addRow(row);
    });

    return workbook;
  }

  static async writeFile(workbook: Workbook, name: string) {
    const path = `${tmpdir()}/${name}_${Date.now().toString(36)}.xlsx`;
    await workbook.xlsx.writeFile(path);
    return path;
  }

  private createRow(opts: { item: any; headers: HeaderRow[], root: string; row?: any }) {
    const { item } = opts;
    let row: any = opts.row ?? {}

    for (const k of opts.headers) {
      let root = opts.root ?? '';

      let value: any = undefined;
      if (k.isNestedPath) {
        value = k.getValueByPath(item, k.path.split(`${root}.`).join(''));
      } else {
        value = k.getValueByPath(item, k.path);
      }

      if (value == null || value == undefined) {
        row[k.path] = '';
        continue;
      }

      root = root == '' ? k.id : `${root}.${k.id}`

      const type = this._getValueType(value);
      if (type == Array && k.sub != null && Array.isArray(k.sub)) {
        const resp = value.flatMap((i: any) =>
          this.createRow({
            item: i,
            headers: k.sub as HeaderRow[],
            root: root,
            row: row,
          }),
        );
        row = { ...row, ...resp };
      } else if (type == Object && k.sub != null) {
        const resp = flatten(
          this.createRow({
            item: value,
            root: root,
            headers: Array.isArray(k.sub) ? k.sub : [k.sub as HeaderRow],
            row: row,
          }),
        );
        row = { ...row, ...resp };
      } else {
        row[k.path] = value ?? '';
      }
    }

    return row;
  }

  /**
   * Recursively transforms raw row header (with nested objects) in a flat map
   *
   */
  private _transformHeadersForExport(headers: HeaderRow[] | HeaderRow): Array<{
    title: string;
    id: string;
    path: string;
    hidden: boolean;
  }> {
    const data = Array.isArray(headers) ? headers : [headers];

    return data.flatMap((i: any) => {
      if (i.sub != null) {
        return this._transformHeadersForExport(i.sub);
      }
      return {
        title: i.title,
        hidden: i.hidden,
        id: i.id,
        path: i.path,
      };
    });
  }

  private generateHeaderRow(opts: {
    root: string;
    data: any[];
    excludeFields: string[];
  }): Array<HeaderRow> {
    const headers: Array<HeaderRow> = [],
      { data, excludeFields } = opts;

    for (const item of data) {
      const keys = sortBy(Object.keys(item));

      for (const k of keys) {
        // Build up a path to the key, e.g. 'foo.bar.baz'
        // If the path exists in the excludeFields array, don't include it in the headers
        const _keyPath = opts.root == '' ? k : `${opts.root}.${k}`;
        if (excludeFields.length > 0) {
          if (
            excludeFields?.indexOf(k) > -1 ||
            excludeFields?.indexOf(_keyPath) > -1
          ) {
            continue;
          }
        }

        const value = item[k];
        const valueType = this._getValueType(value);

        let header = new HeaderRow();

        const existingHeaderIndex = headers.findIndex((i) => i.path == _keyPath);
        if (existingHeaderIndex > -1) {
          header = headers[existingHeaderIndex];
        } else {
          header.id = k;
          header.path = _keyPath;
          header.title = startCase(k);
          header.type = valueType;
          header.hidden = valueType == Object || valueType == Array || valueType == null;
        }


        if (valueType == Object) {
          // TODO; chang data type to set
          const subHeadings = this.generateHeaderRow({
            root: _keyPath,
            data: [value],
            excludeFields: excludeFields,
          });
          if (header.sub != null && Array.isArray(header.sub ?? [])) {
            subHeadings.push(...(header.sub as HeaderRow[]));
          } else if (header.sub != null) {
            subHeadings.push(header.sub as HeaderRow);
          }
          header.sub = unionBy(subHeadings, (h) => h.id).map((v, index) => {
            v.title = `${header.title} ${index + 1} - ${v.title}`;
            return v;
          });
        } else if (valueType == Array) {
          const subCount = ((header.sub as HeaderRow[]) ?? []).length;

          if (header.sub == null || value.length > subCount) {
            header.sub = this.generateHeaderRow({
              root: _keyPath,
              data: value,
              excludeFields: opts.excludeFields,
            }).map((v, index) => {
              v.title = `${header.title} ${index + 1} - ${v.title}`;
              return v;
            });
          }
        }

        if (existingHeaderIndex > -1) {
          headers[existingHeaderIndex] = header;
        } else {
          headers.push(header);
        }
      }
    }

    return headers;
  }

  private _getValueType(value: any) {
    if (value === null || value === undefined) {
      return undefined;
    }

    if (Array.isArray(value)) {
      return Array;
    }

    if (value instanceof Date) {
      return Date;
    }
    // Convert mongoId to string
    // if (isValidObjectId(value)) {
    //   return String;
    // }
    if (typeof value == 'string') {
      return String;
    }
    if (typeof value == 'number') {
      return Number;
    }
    if (typeof value == 'boolean') {
      return Boolean;
    }
    if (typeof value == 'object') {
      return Object;
    }

    return undefined;
  }
}
