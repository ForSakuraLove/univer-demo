import { ComponentManager, IMenuService, MenuGroup, MenuItemType, MenuPosition, } from "@univerjs/ui";
import { IAccessor, Inject, Injector } from "@wendellhu/redi";
import { FolderSingle } from '@univerjs/icons';
import * as UniverJS from "@univerjs/core";
import * as ExcelJS from 'exceljs';

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
        const sheetMap = univerWorkbook.getWorksheets()
        sheetMap.forEach(sheet => {
          // 遍历行
          for (let row = 0; row < sheet.getRowCount(); row++) {
            // 遍历列
            for (let col = 0; col < sheet.getColumnCount(); col++) {
              const cell = sheet.getCell(row, col);
              if (cell) {
                const cellAddress = String.fromCharCode('A'.charCodeAt(0) + col) + (row + 1);
              }
            }
          }
        })
      },
    };
    this.commandService.registerCommand(command);
  }
}

export default ExportExcelButtonPlugin