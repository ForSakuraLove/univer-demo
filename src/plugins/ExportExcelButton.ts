import { ComponentManager, IMenuService, MenuGroup, MenuItemType, MenuPosition, } from "@univerjs/ui";
import { IAccessor, Inject, Injector } from "@wendellhu/redi";
import { FolderSingle } from '@univerjs/icons';
import * as UniverJS from "@univerjs/core";
import * as ExcelJS from 'exceljs';

//格式化时间
const formatDate = (date: Date) => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

//下载文件
const download_file = (buffer: ExcelJS.Buffer, fileName: string) => {
  let fileURL = window.URL.createObjectURL(new Blob([buffer]));
  let fileLink = document.createElement("a");
  fileLink.href = fileURL;
  fileLink.setAttribute("download", fileName);
  document.body.appendChild(fileLink);
  fileLink.click();
}

//下面是导出的函数
const excelExport = async (univerWorkbook: UniverJS.Workbook) => {
  const workbook = new ExcelJS.Workbook();
  const sheetMap = univerWorkbook.getWorksheets()
  sheetMap.forEach(sheet => {
    const worksheet = workbook.addWorksheet(sheet.getName());
    // 遍历行
    for (let row = 0; row < sheet.getRowCount(); row++) {
      const rowData: any[] = [];
      // 遍历列
      for (let col = 0; col < sheet.getColumnCount(); col++) {
        const cell = sheet.getCell(row, col);
        if (cell) {
          rowData[col] = cell.v
        }
      }
      worksheet.addRow(rowData);
    }
  })

  // 写入文件
  const buffer = await workbook.xlsx.writeBuffer();
  //下载文件
  download_file(buffer, formatDate(new Date()) + "导出excel.xlsx");
}

/**
 * Export Excel Button Plugin
 * A simple Plugin example, show how to write a plugin.
 */
class ExportExcelButtonPlugin extends UniverJS.Plugin {
  constructor(
    // inject injector, required
    @Inject(Injector) override readonly _injector: Injector,
    // inject menu service, to add toolbar button
    @Inject(IMenuService) private menuService: IMenuService,
    // inject command service, to register command handler
    @Inject(UniverJS.ICommandService) private readonly commandService: UniverJS.ICommandService,
    // inject component manager, to register icon component
    @Inject(ComponentManager) private readonly componentManager: ComponentManager,
  ) {
    super('export-excel-plugin') // plugin id
  }

  /** Plugin onStarting lifecycle */
  onStarting() {
    this.componentManager.register("FolderSingle", FolderSingle);
    const buttonId = 'export-excel-plugin'
    const menuItem = {
      id: buttonId,
      title: "Export Excel",
      tooltip: "Export Excel",
      icon: "FolderSingle", // icon name
      type: MenuItemType.BUTTON,
      group: MenuGroup.CONTEXT_MENU_DATA,
      positions: [MenuPosition.TOOLBAR_START],
    };
    this.menuService.addMenuItem(menuItem);

    const command = {
      type: UniverJS.CommandType.OPERATION,
      id: buttonId,
      handler: (accessor: IAccessor) => {
        // inject univer instance service
        const univer = accessor.get(UniverJS.IUniverInstanceService);
        const univerWorkbook = univer.getCurrentUniverSheetInstance()
        excelExport(univerWorkbook);
      },
    };
    this.commandService.registerCommand(command);
  }
}

export default ExportExcelButtonPlugin