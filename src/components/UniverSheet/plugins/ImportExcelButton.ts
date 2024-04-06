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
  Workbook,
} from "@univerjs/core";
import { SetRangeValuesCommand } from "@univerjs/sheets";
import { IAccessor, Inject, Injector } from "@wendellhu/redi";
import { FolderSingle } from '@univerjs/icons';
import * as XLSX from 'xlsx';
import { useState } from 'react';


interface President {
  id: string;
  saleNum: number;
}

/* the component state is an array of presidents */
// const [pres, setPres] = useState<President[]>([]);

const waitUserSelectExcelFile = (
  onSelect: (workbook: Workbook) => void,
) => {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".xls, .xlsx";

  input.click();

  input.onchange = () => {
    const file = input.files?.[0];
    if (!file) return;
    console.log(file)
    const reader = new FileReader();

    // reader.onload = () => {
    //   const text = reader.result;
    //   if (typeof text !== "string") return;

    //   // tip: use npm package to parse excel
    //   const rows = text.split(/\r\n|\n/);
    //   const data = rows.map((line) => line.split(","));

    //   const colsCount = data.reduce((max, row) => Math.max(max, row.length), 0);

    //   onSelect({
    //     data,
    //     colsCount,
    //     rowsCount: data.length,
    //   });
    // };
    reader.readAsArrayBuffer(file);
    reader.onload = () => {
      const data = new Uint8Array(reader.result);
      const workbook = XLSX.read(data, { type: 'array' });
      console.log(workbook);

      // 获取第一个工作表的名称
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // 将工作表转换为 JSON 格式
      const excelData = XLSX.utils.sheet_to_json(worksheet);
      console.log(excelData);
      onSelect(workbook);
    };
  };
};

/**
 * parse excel to univer data
 * @param excel
 * @returns { v: string }[][]
 */
const parseExcelUniverData = (excel: string[][]): ICellData[][] => {
  return excel.map((row) => {
    return row.map((cell) => {
      return {
        v: cell || "",
      };
    });
  });
};
function isBinary(str: any) {
  return /^[01]+$/.test(str);
}


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
        const commandService = accessor.get(ICommandService);
        // get current sheet
        const sheet = univer.getCurrentUniverSheetInstance().getActiveSheet();
        // wait user select excel file
        waitUserSelectExcelFile((workbook) => {
          const sheetNames = workbook.getSheetsName()
          const sheets = workbook.getSheets()
          const firstName = sheetNames[0]
          const firstSheet = sheets[0]
          // 将工作表转换为 JSON 格式
          const excelData = XLSX.utils.sheet_to_json(firstSheet);
          console.log(firstSheet);
          console.log(excelData);
          // set sheet size
          // sheet.setColumnCount(colsCount);
          // sheet.setRowCount(rowsCount);

          // set sheet data
          commandService.executeCommand(SetRangeValuesCommand.id, {
            range: {
              // startColumn: 0, // start column index
              // startRow: 0, // start row index
              // endColumn: colsCount - 1, // end column index
              // endRow: rowsCount - 1, // end row index
            },
            // value: parseExcelUniverData(data),
          });
        });
        return true;
      },
    };
    this.commandService.registerCommand(command);
  }
}

export default ImportExcelButtonPlugin