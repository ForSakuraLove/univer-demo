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
} from "@univerjs/core";
import { SetRangeValuesCommand } from "@univerjs/sheets";
import { IAccessor, Inject, Injector } from "@wendellhu/redi";
import { FolderSingle } from '@univerjs/icons';
import * as XLSX from 'xlsx';

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
      if (reader.result instanceof ArrayBuffer) {
        const data = new Uint8Array(reader.result);
        const workbook = XLSX.read(data, { type: 'array' });

        if (workbook.SheetNames.length > 0) {
          onSelect(workbook);
        } else {
          console.error('Workbook does not contain any sheets.');
        }
      } else {
        console.error('Reader result is not an ArrayBuffer.');
      }
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
        waitUserSelectExcelFile((workbook: XLSX.WorkBook) => {
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          // 将工作表转换为 JSON 格式
          // const excelData = XLSX.utils.sheet_to_json(worksheet, { blankrows: true });
          // 将工作表转换为 JSON 格式，并获取标题行和数据
          const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });


          // 获取列数（以第一行为准）
          const numColumns: number = jsonData[0].length;

          // 初始化二维字符串数组
          const twoDimensionalArray: string[][] = [];

          // 遍历一维数组，并将每个元素转换为包含单个字符串的数组，并填充到二维字符串数组中
          jsonData.forEach(row => {
            const newRow: string[] = [];
            for (let i = 0; i < numColumns; i++) {
              newRow.push(String(row[i] || '')); // 如果元素为null或undefined，则用空字符串替代
            }
            twoDimensionalArray.push(newRow);
          });

          console.log(twoDimensionalArray);
          // set sheet size
          sheet.setColumnCount(2);
          sheet.setRowCount(jsonData.length);


          // set sheet data
          commandService.executeCommand(SetRangeValuesCommand.id, {
            range: {
              startColumn: 0, // start column index
              startRow: 0, // start row index
              endColumn: 2, // end column index
              endRow: 11, // end row index
            },
            value: parseExcelUniverData(twoDimensionalArray),
          });
        });
        return true;
      },
    };
    this.commandService.registerCommand(command);
  }
}

export default ImportExcelButtonPlugin