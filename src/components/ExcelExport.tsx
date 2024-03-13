import { Children, cloneElement, useEffect, useMemo, useState } from "react";
import { IExcelWorkerProps } from "../models";
import { isEmpty } from "lodash";
import ExcelWorker from "../workers/excel?worker&inline";

interface IExcelExport extends IExcelWorkerProps {
  fileName: string;
  children: JSX.Element;
  isConvertDataBeforeExport?: boolean;
  onLoading: () => void;
  onSuccess: () => void;
  onEmpty: () => void;
}

const ExcelExport = (props: IExcelExport) => {
  const {
    data,
    columns,
    config,
    fileName,
    mergedFieldCondition,
    children,
    isConvertDataBeforeExport = false,
    onLoading,
    onSuccess,
    onEmpty,
  } = props;
  const [initialized, setInitialized] = useState(false);
  const excelWorker = useMemo(() => new ExcelWorker(), []);
  const enhancedChildren = Children.map(children, (child) =>
    cloneElement(child, {
      onClick: () => {
        !!child.props?.onClick && child.props.onClick();
        if (!isConvertDataBeforeExport) {
          handleExportExcel();
        }
      },
    })
  );

  const handleExportExcel = () => {
    if (!ExcelWorker) return;
    if (isEmpty(data)) {
      onEmpty();
      return;
    }

    onLoading();
    setTimeout(() => {
      const excelWorkerProps: IExcelWorkerProps = {
        data,
        columns,
        config,
        mergedFieldCondition,
      };
      excelWorker.postMessage(excelWorkerProps);
    }, 100);
  };

  useEffect(() => {
    if (ExcelWorker) {
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

  useEffect(() => {
    if (!initialized) {
      setInitialized(true);
      return;
    }
    if (!isConvertDataBeforeExport) {
      return;
    }

    handleExportExcel();
  }, [data]);

  return enhancedChildren;
};

export default ExcelExport;
