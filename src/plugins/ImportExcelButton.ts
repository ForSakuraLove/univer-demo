import {
  ComponentManager,
  IMenuService,
  MenuGroup,
  MenuItemType,
  MenuPosition,
} from "@univerjs/ui";
import {
  CommandType,
  ICellData,
  ICommandService,
  IUniverInstanceService,
  Plugin,
  IRange,
  IRowData,
  IColumnData,
  IWorksheetData,
  BooleanNumber,
  // SheetTypes,
  IFreeze,
  IObjectMatrixPrimitiveType,
  IObjectArrayPrimitiveType,
} from "@univerjs/core";
import { IAccessor, Inject, Injector } from "@wendellhu/redi";
import { FolderSingle } from '@univerjs/icons';
import * as ExcelJS from 'exceljs';



const waitUserSelectExcelFile = (
  onSelect: (workbook: ExcelJS.Workbook) => void,
) => {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".xls, .xlsx";

  input.click();

  input.onchange = () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    const workbook = new ExcelJS.Workbook();
    reader.onload = async () => {
      if (reader.result instanceof ArrayBuffer) {
        const data = new Uint8Array(reader.result);
        await workbook.xlsx.load(data);
        onSelect(workbook);
      }
      else {
        console.error('Reader result is not an ArrayBuffer.');
      }
    };
  };
};
/**
 * Parse Excel worksheet to extract relevant information.
 * @param sheet The Excel worksheet to parse.
 * @returns An object containing information about the worksheet.
 */
// const parseExcelUniverSheetInfo = (sheet: XLSX.WorkSheet, sheetName: string): IWorksheetData => {
const parseExcelUniverSheetInfo = (sheet: ExcelJS.Worksheet): IWorksheetData => {
  const sheetId = sheet.name;
  const name = sheet.name;
  // const type = SheetTypes.GRID;
  const rowCount = sheet.rowCount + 100;
  const columnCount = sheet.columnCount + 100;
  const defaultColumnWidth = 93;
  const defaultRowHeight = 27;
  const scrollTop = 200;
  const scrollLeft = 100;
  const selections = ['A2'];
  const hidden = BooleanNumber.FALSE;
  // const status = 1;
  const showGridlines = 1;
  const rowHeaderWidth = 46;
  const columnHeaderHeight = 20;
  const rightToLeft = BooleanNumber.FALSE;
  const zoomRatio = 1
  const freeze: IFreeze = {
    xSplit: 0, // 水平方向分割的位置
    ySplit: 0, // 垂直方向分割的位置
    startRow: 0, // 冻结区域左上角单元格的行索引，设置为0表示从第一行开始冻结
    startColumn: 0, // 冻结区域左上角单元格的列索引
  };
  const mergeDataBefore: IRange[] = []; // Merged cells
  const cellData: IObjectMatrixPrimitiveType<ICellData> = {}; // Cell data
  const rowData: IObjectArrayPrimitiveType<Partial<IRowData>> = {}; // Row data
  const columnData: IObjectArrayPrimitiveType<Partial<IColumnData>> = {}; // Column data
  sheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (cell.isMerged) {
        // 如果当前单元格是合并单元格，则将其范围添加到 mergeData 数组中
        const masterCell = cell.master;
        mergeDataBefore.push({
          startRow: masterCell.fullAddress.row - 1,
          startColumn: masterCell.fullAddress.col - 1,
          endRow: cell.fullAddress.row - 1,
          endColumn: cell.fullAddress.col - 1
        });
      }
    });
  });

  // 创建一个对象，以便根据起始行和列的组合查找合并范围
  const mergedRangesMap: { [key: string]: IRange } = {};

  // 遍历合并范围数组并整理数据
  mergeDataBefore.forEach(range => {
    const key = `${range.startRow}-${range.startColumn}`;
    if (!mergedRangesMap[key]) {
      mergedRangesMap[key] = range;
    } else {
      // 更新现有范围的结束行和列
      mergedRangesMap[key].endRow = Math.max(mergedRangesMap[key].endRow, range.endRow);
      mergedRangesMap[key].endColumn = Math.max(mergedRangesMap[key].endColumn, range.endColumn);
    }
  });

  // 将整理后的合并范围转换为数组
  const mergeData: IRange[] = Object.values(mergedRangesMap);

  for (let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex++) {
    const row = sheet.getRow(rowIndex)
    for (let colIndex = 1; colIndex <= sheet.columnCount; colIndex++) {
      const cell = row.getCell(colIndex)
      cellData[rowIndex - 1] = cellData[rowIndex - 1] || {}
      rowData[rowIndex - 1] = rowData[rowIndex - 1] || {};
      columnData[colIndex - 1] = columnData[colIndex - 1] || {};
      console.log(cell.$col$row+':')
      if(cell.font){
        if(cell.font.color){
          cellData[rowIndex - 1][colIndex - 1]
        }
      }
      if (cell.value) {
        if (cell.isMerged && cell !== cell.master) {
          cellData[rowIndex - 1][colIndex - 1] = {};
          rowData[rowIndex - 1][colIndex - 1] = {};
          columnData[colIndex - 1][rowIndex - 1] = {};
        } else {
          cellData[rowIndex - 1][colIndex - 1] = { v: cell.value };
          rowData[rowIndex - 1][colIndex - 1] = { v: cell.value };
          columnData[colIndex - 1][rowIndex - 1] = { v: cell.value };

        }
      } else {
        cellData[rowIndex - 1][colIndex - 1] = {};
        rowData[rowIndex - 1][colIndex - 1] = {};
        columnData[colIndex - 1][rowIndex - 1] = {};
      }
    }
  }
  const sheetData: IWorksheetData = {
    /**
   * Id of the worksheet. This should be unique and immutable across the lifecycle of the worksheet.
   */
    id: sheetId.toString(),
    /** Name of the sheet. */
    name: name,
    tabColor: 'white',
    /**
     * Determine whether the sheet is hidden.
     *
     * @remarks
     * See {@link BooleanNumber| the BooleanNumber enum} for more details.
     *
     * @defaultValue `BooleanNumber.FALSE`
     */
    hidden: hidden,
    freeze: freeze,
    rowCount: rowCount,
    columnCount: columnCount,
    zoomRatio: zoomRatio,
    scrollTop: scrollTop,
    scrollLeft: scrollLeft,
    defaultColumnWidth: defaultColumnWidth,
    defaultRowHeight: defaultRowHeight,
    /** All merged cells in this worksheet. */
    mergeData: mergeData,
    /** A matrix storing cell contents by row and column index. */
    cellData: cellData,
    rowData: rowData,
    columnData: columnData,
    rowHeader: {
      width: rowHeaderWidth,
      hidden: hidden,
    },
    columnHeader: {
      height: columnHeaderHeight,
      hidden: hidden,
    },
    showGridlines: showGridlines,
    /** @deprecated */
    selections: selections,
    rightToLeft: rightToLeft,
  };
  return sheetData;
};


/**
 * Import Excel Button Plugin
 * A simple Plugin example, show how to write a plugin.
 */
class ImportExcelButtonPlugin extends Plugin {
  private static onImportExcelCallback?: (data: any) => void;
  constructor(
    // inject injector, required
    @Inject(Injector) override readonly _injector: Injector,
    // inject menu service, to add toolbar button
    @Inject(IMenuService) private menuService: IMenuService,
    // inject command service, to register command handler
    @Inject(ICommandService) private readonly commandService: ICommandService,
    // inject component manager, to register icon component
    @Inject(ComponentManager) private readonly componentManager: ComponentManager,
  ) {
    super('import-excel-plugin') // plugin id
  }
  // 接收函数回调
  static setOnImportExcelCallback(callback: (data: any) => void) {
    ImportExcelButtonPlugin.onImportExcelCallback = callback;
  }

  /** Plugin onStarting lifecycle */
  onStarting() {
    this.componentManager.register("FolderSingle", FolderSingle);
    const buttonId = 'import-excel-plugin'
    const menuItem = {
      id: buttonId,
      title: "Import Excel",
      tooltip: "Import Excel",
      icon: "FolderSingle", // icon name
      type: MenuItemType.BUTTON,
      group: MenuGroup.CONTEXT_MENU_DATA,
      positions: [MenuPosition.TOOLBAR_START],
    };
    this.menuService.addMenuItem(menuItem);

    const command = {
      type: CommandType.OPERATION,
      id: buttonId,
      handler: (accessor: IAccessor) => {
        // inject univer instance service
        const univer = accessor.get(IUniverInstanceService);
        const univerWorkbook = univer.getCurrentUniverSheetInstance()
        const sheetMap = univerWorkbook.getWorksheets()
        sheetMap.forEach(sheet => {
          univerWorkbook.removeSheet(sheet.getSheetId())
        })
        waitUserSelectExcelFile((workbook: ExcelJS.Workbook) => {
          // 处理 Excel 数据
          workbook.eachSheet((worksheet, sheetId) => {
            const sheetInfo: IWorksheetData = parseExcelUniverSheetInfo(worksheet);
            univerWorkbook.addWorksheet(worksheet.name, sheetId, sheetInfo)
          });
          const univeData = univerWorkbook.getSnapshot()
          if (ImportExcelButtonPlugin.onImportExcelCallback) {
            ImportExcelButtonPlugin.onImportExcelCallback(univeData);
          } else {
            console.error("onImportExcelCallback is not defined");
          }
        });
        return true;
      },
    };
    this.commandService.registerCommand(command);
  }
}

export default ImportExcelButtonPlugin