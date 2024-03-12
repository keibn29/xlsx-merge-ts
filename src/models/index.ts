interface IColumn {
  id: number;
  field: string;
  title: string;
  children?: IHeaderColumn[];
  signNumber?: boolean;
  textAlign?: "left" | "center" | "right";
  rowSpan?: number;
  colSpan?: number;
  headerCell?: any;
  cell?: any;
  locked?: boolean;
  icon?: "grid" | "refresh-sm" | "clock" | "delete" | "image-resize";
  type?: "currency" | "date" | "general";
  maxFraction?: number;
  minFraction?: number;
}

export interface IHeaderColumn extends IColumn {
  width?: number;
  excelAlign?: string;
  excelWidth?: number;
  excelFraction?: number;
  isMerge?: boolean;
}

export interface ICellRowSpanConfig {
  key: number;
  distance: number;
}

export interface IExcelWorkerProps {
  data: any[];
  columns: IHeaderColumn[];
  mergedFieldCondition?: string;
}
