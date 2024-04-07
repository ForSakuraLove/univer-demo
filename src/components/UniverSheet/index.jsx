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
import ImportExcelButtonPlugin from '../../plugins/ImportExcelButton';
import ExportExcelButtonPlugin from '../../plugins/ExportExcelButton';
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
    univer.registerPlugin(ExportExcelButtonPlugin);
    // create workbook instance
    workbookRef.current = univer.createUniverSheet(univeData);
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
