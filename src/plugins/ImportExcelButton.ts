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
  SheetTypes,
  IFreeze,
  IObjectMatrixPrimitiveType,
  IObjectArrayPrimitiveType,
} from "@univerjs/core";
import { IAccessor, Inject, Injector } from "@wendellhu/redi";
import { FolderSingle } from '@univerjs/icons';
import * as XLSX from 'xlsx';
import { SetRangeValuesCommand } from "@univerjs/sheets";



const waitUserSelectExcelFile = (
  onSelect: (workbook: XLSX.WorkBook) => void,
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
    reader.onload = () => {
      if (reader.result instanceof ArrayBuffer) {
        const data = new Uint8Array(reader.result);
        const workbook = XLSX.read(data, { type: 'array' });
        if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
          console.error('Workbook does not contain any sheets or is invalid.');
          return;
        }
        onSelect(workbook);
      } else {
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
const parseExcelUniverSheetInfo = (sheet: XLSX.WorkSheet, sheetName: string): IWorksheetData => {
  const sheetId = sheetName;
  // const type = SheetTypes.GRID;
  const name = sheetName;
  const rowCount = 1000;
  const columnCount = 20;
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
  const mergeData: IRange[] = []; // Merged cells
  const cellData: IObjectMatrixPrimitiveType<ICellData> = {}; // Cell data
  const rowData: IObjectArrayPrimitiveType<Partial<IRowData>> = {}; // Row data
  const columnData: IObjectArrayPrimitiveType<Partial<IColumnData>> = {}; // Column data

  if (sheet['!ref']) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex++) {
      for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        const cell = sheet[cellAddress];
        // Extract merged cell information
        if (sheet['!merges']) {
          const mergedCell = sheet['!merges'].find(merge => merge.s.r === rowIndex && merge.s.c === colIndex);
          if (mergedCell) {
            mergeData.push({
              startRow: mergedCell.s.r,
              startColumn: mergedCell.s.c,
              endRow: mergedCell.e.r,
              endColumn: mergedCell.e.c,
            });
          }
        }
        // Extract cell data
        cellData[rowIndex] = cellData[rowIndex] || {};
        cellData[rowIndex][colIndex] = { v: cell?.v };

        // Extract row data
        rowData[rowIndex] = {};

        // Extract column data
        columnData[colIndex] = {};
      }
    }
  }

  const sheetData: IWorksheetData = {
    /**
   * Id of the worksheet. This should be unique and immutable across the lifecycle of the worksheet.
   */
    id: sheetId,
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
        const commandService = accessor.get(ICommandService);
        // wait user select excel file
        waitUserSelectExcelFile((workbook: XLSX.WorkBook) => {
          let sheetIndex = 0;
          Object.keys(workbook.Sheets).forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const sheetInfo: IWorksheetData = parseExcelUniverSheetInfo(sheet, sheetName);
            univerWorkbook.addWorksheet(sheetName, sheetIndex, sheetInfo)
            sheetIndex++;
          });
          commandService.executeCommand(SetRangeValuesCommand.id, {
            range: {
              startColumn: 0,  // start column index
              startRow: 0, // start row index
              endColumn: 100, // end column index
              endRow: 100,  // end row index
            },
          });
        });
        return true;
      },
    };
    this.commandService.registerCommand(command);
  }
}

export default ImportExcelButtonPlugin