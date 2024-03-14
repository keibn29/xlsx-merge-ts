import { cloneDeep } from "lodash";
import { CellObject, ColInfo, WorkSheet, utils, write } from "xlsx-js-style";
import {
  ICellRowSpanConfig,
  IExcelConfig,
  IExcelWorkerProps,
  IHeaderColumn,
} from "../models";

const CELL_ROW_SPAN_FIELD = "cellRowSpan";
const SINGLE_BORDER_STYLE = { style: "thin" };
const CELL_BORDER_STYLE = {
  top: SINGLE_BORDER_STYLE,
  right: SINGLE_BORDER_STYLE,
  bottom: SINGLE_BORDER_STYLE,
  left: SINGLE_BORDER_STYLE,
};
const HEADER_CELL_STYLE = {
  alignment: { vertical: "center", horizontal: "center", wrapText: true },
  fill: { fgColor: { rgb: "7a7a7a" } },
  font: { name: "Malgun Gothic", sz: 11, color: { rgb: "ffffff" }, bold: true },
  border: CELL_BORDER_STYLE,
};
const WPX_WIDTH_DEFAULT = 80;
const WCH_WIDTH_DEFAULT = 14;

const cellRowSpanColumn: IHeaderColumn = {
  id: 0, // id is not important
  title: "",
  field: CELL_ROW_SPAN_FIELD,
};

self.onmessage = (evt: MessageEvent) => {
  const url = generateExcelExportUrl(evt.data);
  postMessage({ url });
};

const generateExcelExportUrl = (props: IExcelWorkerProps) => {
  const dataExcel = convertDataToExcel(props);
  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, dataExcel, "sheet1");
  const excelBuffer = write(workbook, {
    bookType: "xlsx",
    type: "array",
    compression: true,
    cellStyles: true,
    cellDates: true,
  });

  const workbookBlob = new Blob([excelBuffer], {
    type: "application/octet-stream",
  });

  return URL.createObjectURL(workbookBlob);
};

const convertDataToExcel = (props: IExcelWorkerProps) => {
  const { data, columns, mergedFieldCondition, config } = props;
  const depthLength = calculateColumnHeaderDepthLength(columns);
  const excelColumns = columns.concat(cellRowSpanColumn);
  const flattedColumns = flatColumns(excelColumns);
  let dataExport = convertDataToAoa(data, flattedColumns, mergedFieldCondition);
  dataExport = formatData(dataExport, flattedColumns, depthLength, config);
  dataExport = addHeaderMappingIntoDataExcel(dataExport, excelColumns);

  const dataExcel: WorkSheet = utils.aoa_to_sheet(dataExport, {
    dateNF: "yyyy-mm-dd",
  });
  dataExcel["!merges"] = generateExcelMergedConfigs(
    dataExport,
    { flattedColumns, excelColumns },
    depthLength,
    config
  );
  dataExcel["!cols"] = generateColumnWidthConfigs(flattedColumns, config);

  return dataExcel;
};

const flatColumns = (columns: IHeaderColumn[]): IHeaderColumn[] =>
  columns.flatMap((item: IHeaderColumn) => {
    if (item.children) return flatColumns(item.children);
    return item;
  });

const calculateColumnHeaderDepthLength = (
  columns: IHeaderColumn[],
  level = 0
) => {
  let depthLength = level;
  columns.forEach((col: IHeaderColumn) => {
    if (col.children && col.children.length > 0) {
      const nestedLevel = calculateColumnHeaderDepthLength(
        col.children,
        level + 1
      );
      if (nestedLevel > depthLength) {
        depthLength = nestedLevel;
      }
    }
  });
  return depthLength;
};

const convertDataToAoa = (
  data: any[],
  columns: IHeaderColumn[],
  mergedFieldCondition?: string
) => {
  const mergedData = mergedFieldCondition
    ? addCellRowSpanIntoDataExport(data, mergedFieldCondition)
    : data;
  const dataExport = mergedData.map((item: any) => {
    const row: CellObject[] = [];
    columns.forEach((col: IHeaderColumn) => {
      const cell: CellObject = {
        t: "s",
        v: item[col.field] ?? "",
      };
      row.push(cell);
    });
    return row;
  });

  return dataExport;
};

const addCellRowSpanIntoDataExport = (data: any, fieldCondition: string) => {
  const cellData = cloneDeep(data);
  let looped = 1;
  for (let i = 0; i < data.length; i += looped) {
    let rowSpan = 1;
    looped = 1;
    for (let j = i + 1; j < data.length; j++) {
      if (
        data[i][fieldCondition] &&
        data[j][fieldCondition] &&
        data[i][fieldCondition] === data[j][fieldCondition]
      ) {
        looped += 1;
        rowSpan += 1;
      } else {
        break;
      }
    }
    let cellRowSpan: number | undefined = rowSpan;
    if (rowSpan === 1 && looped === 1) cellRowSpan = 1;
    if (rowSpan === 1 && looped !== 1) cellRowSpan = undefined;
    cellData[i][CELL_ROW_SPAN_FIELD] = cellRowSpan;
  }

  return cellData;
};

const formatData = (
  data: any,
  columns: IHeaderColumn[],
  depthLength: number,
  config: IExcelConfig
) => {
  const emptyRow: CellObject[] = [];
  const formattedData = data.map((row: CellObject[], rowIdx: number) => {
    const cellRowSpan = row[row.length - 1].v;
    columns.forEach((col: IHeaderColumn, colIdx: number) => {
      const cellStyle = {
        alignment: {
          vertical: "center",
          horizontal: col[config.alignKey] ?? "center",
          wrapText: true,
        },
        font: { name: "Malgun Gothic", sz: 11 },
        border: CELL_BORDER_STYLE,
      };
      row[colIdx].s = cellStyle;
      row[colIdx] = formatCell(col, row[colIdx], config.fractionKey);
      if (col[config.mergedKey] && !cellRowSpan) row[colIdx].v = 0;

      if (rowIdx === 0) {
        emptyRow.push({ t: "s", v: "", s: HEADER_CELL_STYLE });
      }
    });

    return row;
  });

  formattedData.unshift(
    ...Array.from({ length: depthLength + 1 }, () => cloneDeep(emptyRow))
  );
  return formattedData;
};

const formatCell = (
  col: IHeaderColumn,
  cell: CellObject,
  fractionKey: string
) => {
  const regex = /^\d+(\.\d+)?\s*%$/;
  let formattedCell = cell;

  if (regex.test(`${formattedCell.v ?? ""}`)) {
    formattedCell = formatPercentCell(cell);
  } else if (col.type === "currency") {
    formattedCell = formatCurrencyCell(cell, col[fractionKey]);
  } else if (col.type === "date") {
    formattedCell.z = "yyyy-mm-dd";
  }

  return formattedCell;
};

const formatPercentCell = (cell: CellObject) => {
  const percentValue = Number(`${cell.v ?? 0}`.replace("%", ""));
  cell.t = "n";
  cell.v = percentValue / 100;
  cell.z = Number.isInteger(percentValue) ? "0%" : "0.00%";

  return cell;
};

const formatCurrencyCell = (cell: CellObject, fraction?: number) => {
  const currencyValue = Number(`${cell.v ?? 0}`.replace(/,/g, ""));
  cell.t = "n";
  cell.v = currencyValue;
  cell.z = Number.isInteger(currencyValue) ? "#,##0" : "#,##0.00";
  cell.s.alignment.horizontal = "right";

  if (fraction) {
    cell.z = `#,##0.${"0".repeat(fraction)}`;
  }

  return cell;
};

const generateColumnWidthConfigs = (
  columns: IHeaderColumn[],
  config: IExcelConfig
) => {
  const {
    widthKey,
    unit = "wpx",
    widthRate = 1,
    widthDefault = unit === "wpx" ? WPX_WIDTH_DEFAULT : WCH_WIDTH_DEFAULT,
  } = config;
  const columnWidthConfigs: ColInfo[] = columns.map((col: IHeaderColumn) => {
    const widthConfig =
      col.field === CELL_ROW_SPAN_FIELD
        ? { hidden: true }
        : {
            [unit]: col[widthKey] ? col[widthKey] * widthRate : widthDefault,
          };

    return widthConfig;
  });

  return columnWidthConfigs;
};

const addHeaderMappingIntoDataExcel = (
  dataExport: any[],
  columns: IHeaderColumn[]
) => {
  let columnIndex = 0;

  const recursive = (col: IHeaderColumn, depth: number = 0) => {
    dataExport[depth][columnIndex].v = col.title.replaceAll("<br/>", " ");

    if (!col.children?.length) {
      columnIndex += 1;
    } else {
      col.children.forEach((subCol: IHeaderColumn) => {
        recursive(subCol, depth + 1);
      });
    }
  };

  columns.forEach((col: IHeaderColumn) => {
    recursive(col);
  });

  return dataExport;
};

const generateExcelMergedConfigs = (
  data: any,
  columns: { flattedColumns: IHeaderColumn[]; excelColumns: IHeaderColumn[] },
  depthLength: number,
  config: IExcelConfig
) => {
  const bodyMergedConfig = generateBodyMergedConfigs(
    data,
    columns.flattedColumns,
    config.mergedKey
  );
  const headerNestedMergeConfigs = generateHeaderNestedMergeConfigs(
    columns.excelColumns,
    depthLength
  );

  return bodyMergedConfig.concat(headerNestedMergeConfigs);
};

const generateHeaderNestedMergeConfigs = (
  columns: IHeaderColumn[],
  depthLength: number
) => {
  const headerNestedMergedConfigs: any = [];
  let columnIndex = 0;

  const recursive = (col: IHeaderColumn, depth: number = 0) => {
    if (!col.children?.length) {
      headerNestedMergedConfigs.push({
        s: { c: columnIndex, r: depth },
        e: { c: columnIndex, r: depthLength },
      });
      columnIndex += 1;
    } else {
      const colSpan = calculateColumnHeaderColSpan(col);
      headerNestedMergedConfigs.push({
        s: { c: columnIndex, r: depth },
        e: { c: columnIndex + colSpan - 1, r: depth },
      });

      col.children.forEach((subCol: IHeaderColumn) => {
        recursive(subCol, depth + 1);
      });
    }
  };

  columns.forEach((col: IHeaderColumn) => {
    recursive(col);
  });

  return headerNestedMergedConfigs;
};

const calculateColumnHeaderColSpan = (column: IHeaderColumn) => {
  let colSpan = 1;
  const calculate = (col: IHeaderColumn) => {
    if (col.children?.length) {
      colSpan += col.children.length - 1;
      col.children.forEach((childCol: IHeaderColumn) => {
        calculate(childCol);
      });
    }
  };

  calculate(column);
  return colSpan;
};

const generateBodyMergedConfigs = (
  data: any[],
  columns: IHeaderColumn[],
  mergedKey: string
) => {
  const mergedConfigs: any[] = [];
  const mergedColumnIndexs = columns
    .map((col: IHeaderColumn, idx: number) => (col[mergedKey] ? idx : null))
    .filter((item: number | null) => !!item || item === 0);
  const cellRowSpanConfigs = generateCellRowSpanConfigs(data);

  mergedColumnIndexs.forEach((mergedIdx) => {
    cellRowSpanConfigs.forEach((item) => {
      const config = {
        s: { r: item.key, c: mergedIdx },
        e: { r: item.key + item.distance, c: mergedIdx },
      };
      mergedConfigs.push(config);
    });
  });

  return mergedConfigs;
};

const generateCellRowSpanConfigs = (data: any) => {
  const result: ICellRowSpanConfig[] = [];
  data.forEach((row: CellObject[], idx: number) => {
    const cellRowSpan = row[row.length - 1].v;
    if (cellRowSpan && typeof cellRowSpan === "number" && cellRowSpan > 1) {
      result.push({ key: idx, distance: cellRowSpan - 1 });
    }
  });

  return result;
};
