import { Children, cloneElement, useEffect, useMemo } from "react";
import { IExcelWorkerProps, IHeaderColumn } from "../models";
import { isEmpty } from "lodash";

interface IProps extends IExcelWorkerProps {
  data: any[];
  columns: IHeaderColumn[];
  fileName: string;
  mergedFieldCondition: string;
  onLoading: () => void;
  onSuccess: () => void;
  onNotify: () => void;
  children: JSX.Element;
}

const ExcelExport = (props: IProps) => {
  const {
    data,
    columns,
    fileName,
    mergedFieldCondition,
    onLoading,
    onSuccess,
    onNotify,
    children,
  } = props;
  const enhancedChildren = Children.map(children, (child) =>
    cloneElement(child, {
      onClick: () => {
        !!child.props?.onClick && child.props.onClick();
        handleExportExcel();
      },
    })
  );
  const excelWorker = useMemo(
    () =>
      new Worker(new URL("../workers/excel.ts", import.meta.url), {
        type: "module",
      }),
    []
  );

  const handleExportExcel = () => {
    console.log("heello");
    if (!window.Worker) return;
    if (isEmpty(data)) {
      onNotify();
      return;
    }

    onLoading();
    setTimeout(() => {
      const excelWorkerProps: IExcelWorkerProps = {
        data,
        columns,
        mergedFieldCondition,
      };
      excelWorker.postMessage(excelWorkerProps);
    }, 100);
  };

  useEffect(() => {
    if (window.Worker) {
      excelWorker.onmessage = (evt: MessageEvent) => {
        const a = document.createElement("a");
        a.download = `${fileName}.xlsx`;
        a.href = evt.data.url;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        onSuccess();
      };
    }
  }, [excelWorker, fileName]);

  return enhancedChildren;
};

export default ExcelExport;
