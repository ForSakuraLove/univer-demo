import '@univerjs/design/lib/index.css';
import '@univerjs/ui/lib/index.css';
import '@univerjs/sheets-ui/lib/index.css';
import '@univerjs/sheets-formula/lib/index.css';
import './index.css';
import {
  Univer,
} from "@univerjs/core";
import { defaultTheme } from '@univerjs/design';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverUIPlugin } from '@univerjs/ui';
import { forwardRef, useEffect, useRef, useState } from 'react';
import { FUniver } from "@univerjs/facade";
import * as XLSX from 'xlsx';
import ImportExcelButtonPlugin from '../../plugins/ImportExcelButton';
import { MY_DATA } from '../../assets/my-data'


// eslint-disable-next-line react/display-name
const UniverSheet = forwardRef(() => {
  const univerRef = useRef(null);
  const workbookRef = useRef(null);
  const containerRef = useRef(null);
  const [univeData, setUniveData] = useState(MY_DATA);

  const handleImportExcel = (data) => {
    // 更新组件状态以触发刷新
    setUniveData(data);
  }

  useEffect(() => {
    // 在组件加载后设置回调函数
    ImportExcelButtonPlugin.setOnImportExcelCallback(handleImportExcel);
  }, []);

  /**
   * Initialize univer instance and workbook instance
   * @param data {IWorkbookData} document see https://univer.work/api/core/interfaces/IWorkbookData.html
   */
  const init = () => {
    if (!containerRef.current) {
      throw Error('container not initialized');
    }
    const univer = new Univer({
      theme: defaultTheme,
    });
    univerRef.current = univer;

    // core plugins
    univer.registerPlugin(UniverRenderEnginePlugin);
    univer.registerPlugin(UniverFormulaEnginePlugin);
    univer.registerPlugin(UniverUIPlugin, {
      container: containerRef.current,
      header: true,
      toolbar: true,
      footer: true,
    });

    // doc plugins
    univer.registerPlugin(UniverDocsPlugin, {
      hasScroll: false,
    });
    univer.registerPlugin(UniverDocsUIPlugin);

    // sheet plugins
    univer.registerPlugin(UniverSheetsPlugin);
    univer.registerPlugin(UniverSheetsUIPlugin);
    univer.registerPlugin(UniverSheetsFormulaPlugin);
    univer.registerPlugin(ImportExcelButtonPlugin);

    // create workbook instance
    workbookRef.current = univer.createUniverSheet(univeData);
    const univerAPI = FUniver.newAPI(univer);
    const activeSheet = univerAPI.getActiveWorkbook().getActiveSheet();
    const range = activeSheet.getRange(0, 0, 7, 2);


    var jsonArray = [];
    var keys = [];
    var currentObj = {};
    range.forEach((row, column, cell) => {
      if (row === 0) {
        keys.push(cell.v);
      } else {
        if (column === 0) {
          currentObj[keys[column]] = cell.v;
        } else if (column === 1) {
          currentObj[keys[column]] = cell.v;
          jsonArray.push(currentObj);
          currentObj = {};
        }
      }
    });
    console.log(jsonArray);

    // const worksheet = XLSX.utils.json_to_sheet(jsonArray);
    // const workbook = XLSX.utils.book_new();
    // XLSX.utils.book_append_sheet(workbook, worksheet, "Dates");
    // XLSX.writeFile(workbook, "Presidents" + new Date() + ".xlsx", { compression: true });
  };

  /**
   * Destroy univer instance and workbook instance
   */
  const destroyUniver = () => {
    univerRef.current?.dispose();
    univerRef.current = null;
    workbookRef.current = null;
  };

  useEffect(() => {
    init();
    return () => {
      destroyUniver();
    };
  }, [univeData]);

  return <div ref={containerRef} className="univer-container" />;
});

export default UniverSheet;
